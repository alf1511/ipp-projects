Imports Newtonsoft.Json
Imports System.IO
Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports ClsFungsi
Imports System.Data
Imports System.Xml
Imports System.Collections.Generic
Imports ReportWeb.ClsWebVer

Public Class ServiceConverter

    'Private Function getIPPartner() As String

    '    Dim request = Context.Request

    '    'Return request.ServerVariables("REMOTE_ADDR")

    '    ' Look for a proxy address first
    '    'Dim ip = request.ServerVariables("HTTP_X_FORWARDED_FOR")
    '    Dim ip = "" & request.ServerVariables("HTTP_CLIENT_IP")

    '    ' If there is no proxy, get the standard remote address
    '    If ip.ToString.Trim = "" Or String.Equals(ip, "unknown", StringComparison.OrdinalIgnoreCase) Then

    '        ip = request.ServerVariables("REMOTE_ADDR")

    '    Else

    '        'extract first IP
    '        Dim index = ip.IndexOf(","c)
    '        If index > 0 Then
    '            ip = ip.Substring(0, index)
    '        End If
    '        'remove port
    '        index = ip.IndexOf(":"c)
    '        If index > 0 Then
    '            ip = ip.Substring(0, index)
    '        End If

    '    End If

    '    Return ip

    'End Function

    Public Class _FulfillAccountList
        'format request milik Fulfillment
        Public Account As String = ""
        Public Name As String = ""
        Public _Alias As String = ""
        Public Type As String = ""
        Public Active_Date As String = ""
        Public Add_Time As String = ""
        Public Add_User As String = ""
        Public Upd_Time As String = ""
        Public Upd_User As String = ""
    End Class

    Public Class _FulfillItemList
        'format request milik Fulfillment
        Public Sku As String = ""
        Public Vendor_Code As String = ""
        Public Supplier_Code As String = ""
        Public Name As String = ""
        Public Active_time As String = ""
        Public Add_User As String = ""
        Public Upd_Time As String = ""
        Public Upd_User As String = ""
    End Class

    Public Class _FulfillSyncAccountReq
        Public data As _FulfillAccountList()
    End Class


    Public Class _FulfillSyncItemReq
        Public data As _FulfillItemList()
    End Class

    Public Class _FulfillmentSyncAccountRsp
        'disamakan dengan format response ExpireTransaction
        Public resultcode As String = "9" 'success = 0
        Public message As String = "failed"
        Public description As String = ""
    End Class

    Public Class _FulfillmentSyncItemRsp
        'disamakan dengan format response ExpireTransaction
        Public resultcode As String = "9" 'success = 0
        Public message As String = "failed"
        Public description As String = ""
    End Class

    Public Sub FulfillmentSyncAccount()
        Dim balikan As String = ""

        Dim Method As String = "FulfillmentSyncAccount"

        Dim User As String = "FulfillmentAccount"

        Dim FulfillParameter As String = ""

        Dim IppParameter As String = ""

        Dim ObjFungsi As New ClsFungsi

        Dim LogCheckPoint As String = "0"


        Try
            'Dim IpPartner As String = getIPPartner()

            Dim APIKey As String = ""
            Try
                APIKey = HttpContext.Current.Request.Headers("Authorization")
            Catch ex As Exception
                APIKey = ""
            End Try

            If APIKey = "" Then
                Dim _ObjError As New _FulfillmentSyncAccountRsp
                _ObjError.message = "failed" & " Authorization required"

                balikan = JsonConvert.SerializeObject(_ObjError)
                GoTo Skip
            End If

            LogCheckPoint = "Auth"


            Using sr As New StreamReader(HttpContext.Current.Request.InputStream)
                FulfillParameter = sr.ReadToEnd
            End Using

            If FulfillParameter = "" Then
                Dim _ObjError As New _FulfillmentSyncAccountRsp
                _ObjError.message = "failed" & " request parameter required"

                balikan = JsonConvert.SerializeObject(_ObjError)
                GoTo Skip
            End If

            Dim _ObjFulfillSyncAccount As New _FulfillSyncAccountReq
            _ObjFulfillSyncAccount = JsonConvert.DeserializeObject(Of _FulfillSyncAccountReq)(FulfillParameter)

            If _ObjFulfillSyncAccount Is Nothing Then
                Dim _ObjError As New _FulfillmentSyncAccountRsp
                _ObjError.message = "failed" & " converting request parameter"

                balikan = JsonConvert.SerializeObject(_ObjError)
                GoTo Skip
            End If

            Dim dataTable As New DataTable()

            'Json to Datatable
            For Each row As _FulfillAccountList In _ObjFulfillSyncAccount.data
                'Dim row As DataRow = dataTable.NewRow()
                dataTable.Rows.add(row)
            Next

            Dim url As String = ConfigurationManager.AppSettings("CoreServiceURL")

            Dim serv As New ServiceFulfillment
            Dim servParam(2) As Object

            'serv.Url = url
            'serv.timeout = ConfigurationManager.AppSettings("timeout")

            If _ObjFulfillSyncAccount Is Nothing Then

                servParam = Nothing
            Else

                servParam(0) = "Fulfillment" 'Username
                servParam(1) = APIKey 'Password

            End If

            If Not servParam Is Nothing Then

                Dim objReturn(1) As String
                Dim dtReturn As DataTable
                Dim tempBalikan As New JsonResultResponse_V21

                'Dim tBalikan As Object() = serv.InputFulfillAccount(AppName, AppVersion, servParam, _FulfillSyncAccountReq)

                'objReturn = ConvertArrayObjectToArrayString_V21(tBalikan)
                'dtReturn = ConvertStringToDatatableWithBarcode_V21(objReturn(2))

                tempBalikan.ResultCode = objReturn(0)
                tempBalikan.Message = objReturn(1)

                'balikan = GetJsonData_V21(dtReturn, tempBalikan)
            Else
                'parameter tidak valid

                balikan = "{"
                balikan &= """resultcode"":""" & "9" & ""","
                balikan &= """message"":""" & "Parameter tidak valid" & ""","
                balikan &= """description"":""" & "" & """"
                balikan &= "}"
            End If
Skip:
        Catch ex As Exception
            balikan = "{"
            balikan &= """resultcode"":""" & "9" & ""","
            balikan &= """message"":""" & ex.Message & ""","
            balikan &= """description"":""" & "" & """"
            balikan &= "}"
        End Try
    End Sub

    Public Sub FulfillmentSyncItem(ByVal json As String)
        Dim balikan As String = ""

        Dim Method As String = "FulfillmentSyncItem"

        Dim User As String = "FulfillmentItem"

        Dim FulfillParameter As String = ""

        Dim IppParameter As String = ""

        Dim ObjFungsi As New ClsFungsi

        Dim LogCheckPoint As String = "0"


        Try
            'Dim IpPartner As String = getIPPartner()

            Dim APIKey As String = ""
            Try
                APIKey = HttpContext.Current.Request.Headers("Authorization")
            Catch ex As Exception
                APIKey = ""
            End Try

            'If APIKey = "" Then
            '    Dim _ObjError As New _FulfillmentSyncItemRsp
            '    _ObjError.message = "failed" & " Authorization required"

            '    balikan = JsonConvert.SerializeObject(_ObjError)
            '    GoTo Skip
            'End If

            LogCheckPoint = "Auth"


            'Using sr As New StreamReader(HttpContext.Current.Request.InputStream)
            '    FulfillParameter = sr.ReadToEnd
            'End Using

            FulfillParameter = json

            If FulfillParameter = "" Then
                Dim _ObjError As New _FulfillmentSyncItemRsp
                _ObjError.message = "failed" & " request parameter required"

                balikan = JsonConvert.SerializeObject(_ObjError)
                GoTo Skip
            End If

            Dim _ObjFulfillSyncItem As New _FulfillSyncItemReq
            _ObjFulfillSyncItem = JsonConvert.DeserializeObject(Of _FulfillSyncItemReq)(FulfillParameter)

            If _ObjFulfillSyncItem Is Nothing Then
                Dim _ObjError As New _FulfillmentSyncAccountRsp
                _ObjError.message = "failed" & " converting request parameter"

                balikan = JsonConvert.SerializeObject(_ObjError)
                GoTo Skip
            End If
Skip:
        Catch ex As Exception

        End Try
    End Sub


    Public Class JsonResultResponse_V21
        Public ResultCode As String
        Public Message As String
        Public Description As String

        Function JsonResultToString_V21() As String

            Return "{" + """RESULTCODE""" + ":" + ResultCode + "," + """MESSAGE""" + ":" + """" + Message + """" + "," + """DESCRIPTION""" + ":" + Description + "}"

        End Function

    End Class

End Class

