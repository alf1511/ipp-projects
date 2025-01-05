Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports Newtonsoft.Json
Imports System.IO
Imports System.Web
Imports ReportWeb.ClsFungsi
Imports System.Data
Imports System.Xml
Imports System.Collections.Generic
Imports ReportWeb.ClsWebVer
Imports ReportWeb.ServiceConverter
Imports ReportWeb.CoreService
Imports System.Web.Script.Services


' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
<ScriptService()>
Public Class ServiceConverter1
    Inherits System.Web.Services.WebService
    Public Class _FulfillAccountList
        'format request milik Fulfillment
        Public Account As String = ""
        Public Name As String = ""
        Public [Alias] As String = ""
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
        Public Active_date As String = ""
        Public Inactive_date As String = ""
    End Class

    Public Class _CategoryPacketList
        Public Dokumen As String = ""
        Public Aksesoris As String = ""
        Public Kosmetik As String = ""
        Public ElektronikBaterai As String = ""
        Public ElektronikTanpaBaterai As String = ""
        Public Pakaian As String = ""
        Public SparepartBaterai As String = ""
        Public SparepartTanpaBaterai As String = ""
        Public MakananMinuman As String = ""
        Public DangerousGoods As String = ""
        Public Perishable As String = ""
        Public LiveAnimal As String = ""
        Public LivePlant As String = ""
    End Class
    Public Class _CategoryPacketReq
        Public data As _CategoryPacketList()
    End Class

    Public Class _FulfillSyncAccountReq
        Public data As _FulfillAccountList()
    End Class

    Public Class _FulfillmentAddItemReq
        Public data As _FulfillItemList()
    End Class
    Public Class _CategoryPacketReqRsp
        'disamakan dengan format response ExpireTransaction
        Public resultcode As String = "9" 'success = 0
        Public message As String = "failed"
        Public description As String = ""
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

    <WebMethod()>
    Public Sub FulfillmentSyncAccount()
        Dim balikan As String = ""

        Dim Method As String = "FulfillmentSyncAccount"

        Dim User As String = "FulfillmentAccount"

        Dim FulfillParameter As String = ""

        Dim IppParameter As String = ""

        Dim ObjFungsi As New ClsFungsi

        Dim LogCheckPoint As String = "0"


        Try
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

            Dim dataJson As String = JsonConvert.SerializeObject(_ObjFulfillSyncAccount.data)
            Dim dataTable As DataTable = JsonConvert.DeserializeObject(Of DataTable)(dataJson)

            Dim dataSet As New DataSet

            dataSet.Tables.Add(dataTable)

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

                Dim tBalikan As Object() = serv.InputFulfillAccount(AppName, AppVersion, servParam, dataSet)

                objReturn = ConvertArrayObjectToArrayString_V21(tBalikan)
                dtReturn = ConvertStringToDatatableWithBarcode_V21(objReturn(2))

                tempBalikan.ResultCode = objReturn(0)
                tempBalikan.Message = objReturn(1)

                balikan = GetJsonData_V21(dtReturn, tempBalikan)
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

    Public Function ConvertArrayObjectToArrayString_V21(ByVal Parameter As Object()) As String()

        On Error Resume Next

        Dim Result As String()
        ReDim Result(Parameter.Length - 1)

        For i As Integer = 0 To Result.Length - 1
            Result(i) = ("" & Parameter(i)).ToString
        Next

        Return Result

    End Function

    Public Function ConvertStringToDatatableWithBarcode_V21(ByVal Param As String) As DataTable
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
        Catch ex As Exception
            dtResult = Nothing
        End Try
        Return dtResult
    End Function

    Public Function GetJsonData_V21(ByVal dtReturn As DataTable, ByVal tempBalikan As JsonResultResponse_V21) As String

        Dim ds As New DataSet
        Dim doc As New XmlDocument()
        Dim json As String
        Dim node As XmlNode
        Dim jsonLength As Int32

        ds.Tables.Add(dtReturn)
        doc.LoadXml(ds.GetXml())
        json = JsonConvert.SerializeObject(doc)
        jsonLength = (json.IndexOf("]") - json.IndexOf("[")) + 1

        If jsonLength = 1 Then 'Melakukan cek karena apabila data hanya 1 tidak akan membentuk array []

            node = doc.FirstChild.FirstChild 'Mengambil node dari object
            json = JsonConvert.SerializeXmlNode(node, 0, True) 'serializing xml node dengan menghikangkan node parent
            If json = "null" Then 'cek apabila array kosong untuk menghilangkan atribut null pada array
                tempBalikan.Description = "[]"
            Else
                tempBalikan.Description = "[" + json + "]"
            End If

        Else

            tempBalikan.Description = json.Substring(json.IndexOf("["), (json.IndexOf("]") - json.IndexOf("[")) + 1)

        End If

        Return tempBalikan.JsonResultToString_V21()
    End Function

    Public Class _FulfillmentRsp
        Public resultcode As String = ""
        Public message As String = ""
        Public description As String = ""
    End Class

    Dim wsAppName As String = "WServiceFulfillment"
    Dim wsAppVersion As String = "24.09.12.00"
    <WebMethod()>
    Public Sub FulfillmentAddItem(ByVal json As String)
        Dim balikan As String = ""

        Dim Method As String = "FulfillmentAddItem"

        'Dim IppParameter As String = ""

        'Dim ObjFungsi As New ClsFungsi

        Dim LogCheckPoint As String = "0"

        Try
            Dim APIUser As String = ""
            Try
                APIUser = HttpContext.Current.Request.Headers("Username")
            Catch ex As Exception
                APIUser = ""
            End Try

            Dim APIKey As String = ""
            Try
                APIKey = HttpContext.Current.Request.Headers("Authorization")
            Catch ex As Exception
                APIKey = ""
            End Try

            'If APIUser = "" Or APIKey = "" Then
            '    Dim _ObjError As New _FulfillmentRsp
            '    _ObjError.message = "failed" & " Authorization required"

            '    balikan = JsonConvert.SerializeObject(_ObjError)
            '    GoTo Skip
            'End If

            LogCheckPoint = "Auth"


            Dim FulfillParameter As String = ""
            'Using sr As New StreamReader(HttpContext.Current.Request.InputStream)
            '    FulfillParameter = sr.ReadToEnd
            'End Using

            FulfillParameter = json

            If FulfillParameter = "" Then
                Dim _ObjError As New _FulfillmentRsp
                _ObjError.message = "failed" & " request parameter required"

                balikan = JsonConvert.SerializeObject(_ObjError)
                GoTo Skip
            End If

            Dim _ObjReq As New _FulfillmentAddItemReq
            _ObjReq = JsonConvert.DeserializeObject(Of _FulfillmentAddItemReq)(FulfillParameter)

            If _ObjReq Is Nothing Then
                Dim _ObjError As New _FulfillmentRsp
                _ObjError.message = "failed" & " converting request parameter"

                balikan = JsonConvert.SerializeObject(_ObjError)
                GoTo Skip
            End If

            Dim dataJson As String = JsonConvert.SerializeObject(_ObjReq.data)
            Dim dataTable As DataTable = JsonConvert.DeserializeObject(Of DataTable)(dataJson)

            Dim dataSet As New DataSet

            dataSet.Tables.Add(dataTable)

            Dim url As String = ConfigurationManager.AppSettings("ServiceFulfillment.ServiceFulfillment")

            'Dim serv As New ServiceFulfillment.ServiceFulfillment
            Dim serv As New ServiceFulfillment
            Dim servParam(1) As Object

            'serv.Url = url
            'serv.Timeout = ConfigurationManager.AppSettings("timeout")

            If _ObjReq Is Nothing Then

                servParam = Nothing
            Else

                servParam(0) = APIUser
                servParam(1) = APIKey

            End If

            If Not servParam Is Nothing Then

                Dim objReturn(1) As String
                Dim dtReturn As DataTable
                Dim tempBalikan As New JsonResultResponse_V21

                Dim tBalikan As Object() = serv.FulfillmentAddItem(wsAppName, wsAppVersion, servParam, dataSet)

                objReturn = ConvertArrayObjectToArrayString_V21(tBalikan)
                dtReturn = ConvertStringToDatatableWithBarcode_V21(objReturn(2))

                tempBalikan.ResultCode = objReturn(0)
                tempBalikan.Message = objReturn(1)

                balikan = GetJsonData_V21(dtReturn, tempBalikan)
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

        Context.Response.AddHeader("Content-Type", "application/json")
        Context.Response.Output.Write(balikan)
        Context.Response.End()

    End Sub

    <WebMethod()>
    Public Sub GetPackageList(ByVal Parameter As String, ByVal Jenis As String)

        'Dim IpPartner As String = getIPPartner()

        Dim ObjFungsi As New ClsFungsi
        Dim balikan As String = ObjFungsi.GetPackageList("" & Parameter, "" & Jenis)

        If ("" & Jenis).ToUpper = "JSON" Then
            Context.Response.AddHeader("Content-Type", "application/json")
        End If
        Context.Response.Output.Write(balikan)
        Context.Response.End()

        'JSON
        '----------------
        '{
        '	"User":"itsd7",
        '	"Password":"777777",
        '	"Category":"KLIKINDOMARET",
        '	"Description":"1"
        '}

        'XML
        '----------------
        '<parameter>
        '  <User>itsd7</User>
        '  <Password>777777</Password>
        '  <Category>KLIKINDOMARET</Category>
        '  <Description>1</Description>
        '</parameter>

    End Sub

    <WebMethod()>
    Public Function HelloWorld() As String
        Return "Hello World"
    End Function

End Class