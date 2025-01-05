'Imports System.Web
'Imports System.Web.Services
'Imports System.Web.Services.Protocols
'Imports System.Data
'Imports MySql.Data.MySqlClient
'Imports Newtonsoft.Json
'Imports System.Collections.Generic
'Imports ReportWeb.ClsWebVer

'<WebService(Namespace:="http://tempuri.org/")>
'<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
'<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
'Public Class ServiceDeliveryMan
'    Inherits System.Web.Services.WebService

'    Private MasterMCon As MySqlConnection
'    Private SqlParam As New Dictionary(Of String, String)

'    Private WService As New LocalCore

'    Private wsUserAgent As String = "INDOPAKET"
'    Private wsUser As String
'    Private wsPassword As String

'    Private CustomAccount3rdPartyDeliveryMan As String
'    Private CustomAccount3rdPartyIndopaketMotor As String
'    Private CustomAccount3rdPartyIndopaketMobil As String
'    Private CustomAccount3rdPartyIndopaketInstantMotor As String
'    Private CustomAccount3rdPartyIndopaketInstantMobil As String

'    Private CustomAccountExpeditionDMSMotor As String
'    Private CustomAccountExpeditionDMSMobil As String

'    'Private mURLOAuth As String = ""
'    'Private mURLPrice As String = ""
'    Private mURLOrder As String = ""
'    Private mURLEditOrder As String = ""
'    'Private mURLTracking As String = ""
'    Private mURLSyncToko As String = ""
'    Private mURLSendDataKaryawan As String = ""
'    Private mContentType As String = ""
'    Private mTimeout As String = ""

'    Private CustomHeaders As String() = Nothing

'    Private CustomAccountKlikApka As String
'    Private CustomAccountKlikIStore As String
'    Private CustomAccountKlikFood As String
'    Private CustomAccountKlikReOrder As String
'    Private CustomAccountServerInventoryDelimanCode As String = "'SID'"

'    Private CustomAccountLazada As String
'    Private CustomAccountLazadaReverse As String

'    Private KeywordIppDoorPickupServiceType As String = "IPPDOORPICKUP" 'disamakan dengan Service.vb

'    Public Sub New()

'        Dim ConSQL As String = "" & ConfigurationManager.AppSettings("ConSQL")

'        If ConSQL <> "" Then
'            'webconfig
'            MasterMCon = New MySqlConnection(ConSQL)
'        Else
'            'default
'            MasterMCon = New MySqlConnection("Server=localhost;Port=3306;Database=iexpress;Uid=root;Pwd=root;")
'        End If

'        'mURLOAuth = ("" & ConfigurationManager.AppSettings("DeliveryManAPIOAuth")).Trim
'        'If mURLOAuth = "" Then
'        '    mURLOAuth = "x"
'        'End If

'        'mURLPrice = ("" & ConfigurationManager.AppSettings("DeliveryManAPIPrice")).Trim
'        'If mURLPrice = "" Then
'        '    mURLPrice = "x"
'        'End If

'        mURLOrder = ("" & ConfigurationManager.AppSettings("DeliveryManAPIOrder")).Trim
'        If mURLOrder = "" Then
'            mURLOrder = "x"
'        End If

'        mURLEditOrder = ("" & ConfigurationManager.AppSettings("DeliveryManAPIEditOrder")).Trim
'        If mURLEditOrder = "" Then
'            mURLEditOrder = "x"
'        End If

'        'mURLTracking = ("" & ConfigurationManager.AppSettings("DeliveryManAPITracking")).Trim
'        'If mURLTracking = "" Then
'        '    mURLTracking = "x"
'        'End If

'        mURLSyncToko = ("" & ConfigurationManager.AppSettings("DeliveryManAPISyncToko")).Trim
'        If mURLSyncToko = "" Then
'            mURLSyncToko = "x"
'        End If

'        mURLSendDataKaryawan = ("" & ConfigurationManager.AppSettings("DeliveryManSendDataKaryawanUrl")).Trim
'        If mURLSendDataKaryawan = "" Then
'            mURLSendDataKaryawan = "x"
'        End If

'        wsUser = ("" & ConfigurationManager.AppSettings("DeliveryManAPIWsUser")).Trim
'        wsPassword = ("" & ConfigurationManager.AppSettings("DeliveryManAPIWsPassword")).Trim
'        If wsUser = "" Or wsPassword = "" Then
'            wsUser = "x"
'            wsPassword = "x"
'        End If

'        mContentType = "application/json"

'        mTimeout = ("" & ConfigurationManager.AppSettings("DeliveryManAPIWsTimeout")).Trim 'dalam detik
'        If mTimeout = "" Then
'            mTimeout = "60"
'        End If
'        'ws.Timeout = mTimeout * 1000

'        CustomAccount3rdPartyDeliveryMan = ("" & ConfigurationManager.AppSettings("DeliveryManAccount3rdParty")).Trim
'        CustomAccount3rdPartyIndopaketMotor = ("" & ConfigurationManager.AppSettings("IndopaketMotorAccount3rdParty")).Trim
'        CustomAccount3rdPartyIndopaketMobil = ("" & ConfigurationManager.AppSettings("IndopaketMobilAccount3rdParty")).Trim
'        CustomAccount3rdPartyIndopaketInstantMotor = ("" & ConfigurationManager.AppSettings("IndopaketInstantMotorAccount3rdParty")).Trim
'        CustomAccount3rdPartyIndopaketInstantMobil = ("" & ConfigurationManager.AppSettings("IndopaketInstantMobilAccount3rdParty")).Trim

'        CustomAccountExpeditionDMSMotor = ("" & ConfigurationManager.AppSettings("DeliveryManMotorAccountExpedition")).Trim
'        CustomAccountExpeditionDMSMobil = ("" & ConfigurationManager.AppSettings("DeliveryManMobilAccountExpedition")).Trim

'        CustomAccountKlikApka = ("" & ConfigurationManager.AppSettings("KlikApkaAccountECo")).Trim
'        CustomAccountKlikIStore = ("" & ConfigurationManager.AppSettings("KlikIStoreAccountECo")).Trim
'        CustomAccountKlikFood = ("" & ConfigurationManager.AppSettings("KlikFoodAccountECo")).Trim
'        CustomAccountKlikReOrder = ("" & ConfigurationManager.AppSettings("KlikReOrderAccountECo")).Trim

'        CustomAccountLazada = ("" & ConfigurationManager.AppSettings("LazadaAccountECo")).Trim
'        CustomAccountLazadaReverse = ("" & ConfigurationManager.AppSettings("LazadaReverseAccountECo")).Trim

'    End Sub


'    <WebMethod()>
'    Public Function ChekConnection() As String
'        Return "OK"
'    End Function

'    Private Function CreateResult() As Object()

'        Dim Result As Object()
'        ReDim Result(3)

'        Result(0) = "9" 'Kode Proses
'        Result(1) = "" 'Pesan hasil proses / pesan error
'        Result(2) = Nothing 'Variable yang dikembalikan

'        Return Result

'    End Function


'    <WebMethod()>
'    Public Function GetOAuth(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object

'        Dim Method As String = "GetOAuth"

'        Dim Result As Object() = CreateResult()
'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString

'            Dim MustReset As Boolean = False
'            Try
'                MustReset = (Param(2).ToString = "1")
'            Catch ex As Exception
'                MustReset = False
'            End Try


'            MCon = MasterMCon.Clone

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            Dim ObjFungsi As New ClsFungsi

'            Dim e As New ClsError

'            Dim ObjSQL As New ClsSQL

'            Dim Access_Token As String = ""

'            If MustReset = False Then
'                'get local oauth
'                Query = "Select OAuth From AutoOrderOAuth"
'                Query &= " Where `User` in (" & CustomAccount3rdPartyDeliveryMan & ")"
'                Try
'                    Access_Token = ObjSQL.ExecScalar(MCon, Query).ToString
'                Catch ex As Exception
'                    Access_Token = ""
'                End Try
'            End If 'If MustReset = False

'            If Access_Token = "" Then

'                Access_Token = wsPassword

'                If Access_Token.Trim = "" Then
'                    Result(1) = "Gagal GetOAuth"
'                    GoTo Skip
'                End If

'                'save local oauth
'                Query = "Insert Into AutoOrderOAuth ("
'                Query &= " `User`, OAuth, UpdTime, UpdUser"
'                Query &= " ) values ("
'                Query &= " '" & CustomAccount3rdPartyDeliveryMan.Replace("'", "") & "','" & Access_Token & "'"
'                Query &= " , now(), '" & "WServiceDeliMan" & "'"
'                Query &= " ) on duplicate key update"
'                Query &= " OAuth = '" & Access_Token & "'"
'                Query &= " , UpdTime = now(), UpdUser = '" & "WServiceDeliMan" & "'"
'                ObjSQL.ExecScalar(MCon, Query)

'            End If 'dari If Access_Token = ""


'            ReDim Result(2)
'            'Result(4) = JsonConvert.SerializeObject(ObjRspOAuth)
'            'Result(3) = JsonConvert.SerializeObject(ObjReqOAuth)
'            Result(2) = Access_Token
'            Result(1) = ""
'            Result(0) = "0"

'Skip:

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(AppName, AppVersion, Method, "", ex, Query)

'            Result(1) = ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Public Class cReqCreateOrder

'        Public awbId As String = "" 'AWB Indopaket
'        Public awbIdIpp As String = "" 'AWB Indopaket tanpa ekor dTracknum
'        Public salesOrderId As String = "" 'Order ID Partner

'        Public serviceType As String = ""
'        Public orderType As String = ""
'        Public channelOrder As String = ""
'        Public vehicleType As String = "" 'MOTOR / MOBIL / TRUCK
'        Public storeCode As String = ""

'        Public senderName As String = "" 'nama aktual pengirim, untuk paket to Door, ketika MTO di MS, bukan informasikan nama toko

'        Public pickupPin As String = ""
'        Public cancelPin As String = ""
'        Public keepPin As String = ""
'        Public returnPin As String = ""

'        Public minDeliveryTime As String = ""
'        Public maxDeliveryTime As String = ""

'        Public minDeliveryTimeOri As String = ""
'        Public maxDeliveryTimeOri As String = ""

'        Public origin As New cReqCreateOrderOri

'        Public recipient As New cReqCreateOrderDst

'        Public cod As Boolean = False
'        Public codAmount As Double = 0
'        Public codPaymentCode As String = ""
'        Public codPaymentBiller As String = ""

'        Public bulky As Boolean = False
'        Public item As cReqCreateOrderBulky()

'        Public doorPickupPaymentStatus As String = "" 'Paid/Unpaid
'        Public doorPickupPaymentValue As Double = 0

'    End Class

'    Public Class cReqCreateOrderOri
'        Public pickupName As String = ""
'        Public pickupPhoneNumber As String = ""

'        Public pickupAddress As String = ""
'        Public pickupLatitude As Double
'        Public pickupLongitude As Double
'        Public pickupNotes As String = ""

'        'Public pickupDriverLocation As Boolean = False
'    End Class

'    Public Class cReqCreateOrderDst
'        Public customerName As String = ""
'        Public customerPhoneNumber As String = ""

'        Public deliveryAddress As String = ""
'        Public deliveryLatitude As Double
'        Public deliveryLongitude As Double
'        Public deliveryNotes As String = ""

'        'Public deliveryDriverLocation As Boolean = False
'    End Class

'    Public Class cReqCreateOrderBulky
'        Public itemCode As String = ""
'        Public itemName As String = ""
'        Public quantity As Integer = 0
'    End Class

'    Public Class cRspCreateOrder
'        '00 = success
'        '03 = bad request
'        'IO1001 = duplicate order
'        'IO1002 = parse error

'        Public status As String = ""
'        Public message As String = ""

'        Public data As New cRspCreateOrderDetail

'        Public timestamp As String = ""
'        Public signature As String = ""
'    End Class

'    Public Class cRspCreateOrderDetail
'        Public orderId As String = ""
'        Public awbId As String = ""
'        Public salesOrderId As String = ""
'        Public transactionDate As String = ""
'    End Class

'    <WebMethod()>
'    Public Function CreateOrder(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object() _
'    , ByVal dsData As DataSet) As Object

'        Dim Method As String = "CreateOrder"

'        Dim Result As Object() = CreateResult()
'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Dim LogKeyword As String = ""

'        Try
'            Dim Password As String = Param(1).ToString
'            Dim InsertAutoOrderTracking As String = "0"
'            Try
'                InsertAutoOrderTracking = Param(2).ToString
'            Catch ex As Exception
'                InsertAutoOrderTracking = "0"
'            End Try

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim sDistance As String = "-1"
'            Dim sPrice As String = "-1"
'            Dim sOrder_Id As String = ""

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            Dim dtService As New DataTable
'            Dim dtShipper As New DataTable
'            Dim dtConsignee As New DataTable
'            'Dim dtPackage As New DataTable
'            Dim dtPIN As New DataTable
'            Dim dtCOD As New DataTable
'            Dim dtReturnBulky As New DataTable

'            For d As Integer = 0 To dsData.Tables.Count - 1
'                Select Case dsData.Tables(d).TableName.ToUpper
'                    Case "SERVICE"
'                        dtService = dsData.Tables(d).Copy

'                    Case "SHIPPER"
'                        dtShipper = dsData.Tables(d).Copy

'                    Case "CONSIGNEE"
'                        dtConsignee = dsData.Tables(d).Copy

'                        'Case "PACKAGE"
'                        '    dtPackage = dsData.Tables(d).Copy

'                    Case "PIN"
'                        dtPIN = dsData.Tables(d).Copy

'                    Case "COD"
'                        dtCOD = dsData.Tables(d).Copy

'                    Case "RETURNBULKY"
'                        dtReturnBulky = dsData.Tables(d).Copy

'                End Select
'            Next

'            If dtService.Rows.Count < 1 Or dtShipper.Rows.Count < 1 _
'            Or dtConsignee.Rows.Count < 1 Or dtPIN.Rows.Count < 1 Then
'                Result(1) = "Data tidak lengkap"
'                GoTo Skip
'            End If


'            Dim ObjSql As New ClsSQL

'            LogKeyword = dtService.Rows(0).Item("OrderNo").ToString

'            Dim ActualTrackNum As String = ""
'            Try
'                ActualTrackNum = dtService.Rows(0).Item("ActualTrackNum").ToString
'            Catch ex As Exception
'                ActualTrackNum = ""
'            End Try

'            Dim TrackNum As String = dtService.Rows(0).Item("TrackNum").ToString
'            Dim OrderNo As String = dtService.Rows(0).Item("OrderNo").ToString
'            Dim ServiceType As String = dtService.Rows(0).Item("ServiceType").ToString

'            Dim OrderType As String = ""
'            Try
'                OrderType = dtService.Rows(0).Item("OrderType").ToString
'            Catch ex As Exception
'                OrderType = ""
'            End Try

'            Dim OrderTypeDetail As String = ""
'            Try
'                OrderTypeDetail = dtService.Rows(0).Item("OrderTypeDetail").ToString.Trim.ToUpper
'            Catch ex As Exception
'                OrderTypeDetail = ""
'            End Try

'            Dim ChannelOrder As String = ""
'            Select Case OrderTypeDetail
'                Case "KLIKINDOGROSIR"
'                    ChannelOrder = "INDOGROSIR"
'                Case "OTHER"
'                    ChannelOrder = "OTHER"
'                Case Else
'                    ChannelOrder = "INDOMARET"
'            End Select

'            'pindah ke Service.AutoOrderDoorPickup
'            'If ServiceType.ToUpper = KeywordIppDoorPickupServiceType.ToUpper Then
'            '    ServiceType = "SAMEDAY"
'            '    OrderType = "PAKET-JEMPUT"
'            '    ChannelOrder = "INDOMARET"
'            'End If

'            If ChannelOrder.ToUpper = "INDOMARET" Then
'                If OrderType.ToUpper.Contains("GROCERY") Or OrderType.ToUpper.Contains("FOOD") Then
'                    If ServiceType.ToUpper.Contains("REGULAR") Then
'                        ServiceType = "REGULAR-SLOT"
'                    End If
'                End If
'            End If

'            Dim VehicleType As String = "X"
'            Try
'                VehicleType = dtService.Rows(0).Item("VehicleType").ToString.Trim.ToUpper
'            Catch ex As Exception
'                VehicleType = "X"
'            End Try

'            Dim StoreCode As String = ""
'            Try
'                StoreCode = dtService.Rows(0).Item("StoreCode").ToString.Trim.ToUpper
'            Catch ex As Exception
'                StoreCode = ""
'            End Try

'            Dim ActualSenderName As String = ""
'            Try
'                ActualSenderName = dtService.Rows(0).Item("ActualSenderName").ToString
'            Catch ex As Exception
'                ActualSenderName = ""
'            End Try

'            Dim DeliveryMinTime As String = dtService.Rows(0).Item("DeliveryMinTime").ToString
'            Dim DeliveryMaxTime As String = dtService.Rows(0).Item("DeliveryMaxTime").ToString

'            Dim DeliveryMinTimeOri As String = dtService.Rows(0).Item("DeliveryMinTimeOri").ToString
'            Dim DeliveryMaxTimeOri As String = dtService.Rows(0).Item("DeliveryMaxTimeOri").ToString

'            Dim IsBulky As Boolean = False
'            Try
'                IsBulky = (dtService.Rows(0).Item("FlgBulky").ToString = "1")
'            Catch ex As Exception
'                IsBulky = False
'            End Try

'            Dim ShPaymentStatus As String = "UNPAID"
'            Try
'                If dtService.Rows(0).Item("PaymentStatus").ToString = "1" Then
'                    ShPaymentStatus = "PAID"
'                End If
'            Catch ex As Exception
'                ShPaymentStatus = "UNPAID"
'            End Try
'            Dim ShPaymentValue As Double = 0
'            Try
'                ShPaymentValue = CDbl(dtService.Rows(0).Item("PaymentValue"))
'            Catch ex As Exception
'                ShPaymentValue = 0
'            End Try


'            Dim ShName As String = dtShipper.Rows(0).Item("Name").ToString
'            Dim ShPhone As String = FormatPhone(dtShipper.Rows(0).Item("Phone").ToString)
'            'Dim ShEmail As String = dtShipper.Rows(0).Item("Email").ToString
'            Dim ShAddress As String = dtShipper.Rows(0).Item("Address").ToString
'            Dim ShPostalCode As String = dtShipper.Rows(0).Item("PostalCode").ToString
'            Dim ShLatitude As String = dtShipper.Rows(0).Item("Latitude").ToString
'            Dim ShLongitude As String = dtShipper.Rows(0).Item("Longitude").ToString
'            Dim ShNotes As String = ""
'            Try
'                ShNotes = dtShipper.Rows(0).Item("Note").ToString
'            Catch ex As Exception
'                ShNotes = ""
'            End Try

'            Dim CoName As String = dtConsignee.Rows(0).Item("Name").ToString
'            Dim CoPhone As String = FormatPhone(dtConsignee.Rows(0).Item("Phone").ToString)
'            'Dim CoEmail As String = dtConsignee.Rows(0).Item("Email").ToString
'            Dim CoAddress As String = dtConsignee.Rows(0).Item("Address").ToString
'            Dim CoPostalCode As String = dtConsignee.Rows(0).Item("PostalCode").ToString
'            Dim CoLatitude As String = dtConsignee.Rows(0).Item("Latitude").ToString
'            Dim CoLongitude As String = dtConsignee.Rows(0).Item("Longitude").ToString
'            Dim CoNotes As String = ""
'            Try
'                CoNotes = dtConsignee.Rows(0).Item("Note").ToString
'            Catch ex As Exception
'                CoNotes = ""
'            End Try
'            Dim CoAddInfo As String = ""
'            Try
'                CoAddInfo = dtConsignee.Rows(0).Item("AddInfo").ToString.Trim.ToUpper
'            Catch ex As Exception
'                CoAddInfo = ""
'            End Try

'            'Dim PckName As String = dtPackage.Rows(0).Item("Name").ToString
'            'Dim PckDesc As String = dtPackage.Rows(0).Item("Description").ToString
'            'Dim PckQty As Integer = CInt(dtPackage.Rows(0).Item("Qty"))
'            'Dim PckValue As Integer = CInt(dtPackage.Rows(0).Item("Value"))
'            'Dim PckWeight As Integer = CInt(dtPackage.Rows(0).Item("Weight"))
'            'Dim PckLength As Integer = CInt(dtPackage.Rows(0).Item("Length"))
'            'Dim PckWidth As Integer = CInt(dtPackage.Rows(0).Item("Width"))
'            'Dim PckHeight As Integer = CInt(dtPackage.Rows(0).Item("Height"))

'            Dim PINPickup As String = dtPIN.Rows(0).Item("Pickup").ToString
'            Dim PINCancel As String = dtPIN.Rows(0).Item("Cancel").ToString
'            Dim PINKeep As String = dtPIN.Rows(0).Item("Keep").ToString
'            Dim PINReturn As String = dtPIN.Rows(0).Item("Return").ToString


'            Dim CODValue As Double = 0
'            Try
'                CODValue = CDbl(dtCOD.Rows(0).Item("CODValue"))
'            Catch ex As Exception
'                CODValue = 0
'            End Try

'            Dim CODPaymentCode As String = ""
'            Try
'                CODPaymentCode = dtCOD.Rows(0).Item("CODPaymentCode").ToString
'            Catch ex As Exception
'                CODPaymentCode = ""
'            End Try

'            Dim CODPaymentBiller As String = ""
'            Try
'                CODPaymentBiller = dtCOD.Rows(0).Item("CODPaymentBiller").ToString
'            Catch ex As Exception
'                CODPaymentBiller = ""
'            End Try

'            'Dim FlgCOD As Boolean = (CODValue > 0 And CODPaymentCode <> "")
'            Dim FlgCOD As Boolean = (CODValue > 0)


'            Dim ObjFungsi As New ClsFungsi


'            e.DebugLog(MCon, wsAppVersion, Method, "WServiceDeliMan", "START", LogKeyword)

'            'Dim Access_Token As String = ""

'            'Dim tResult As Object() = GetOAuth(AppName, AppVersion, Param)
'            'If tResult(0).ToString = "0" Then
'            '    Access_Token = tResult(2).ToString
'            'Else
'            '    Result(1) = tResult(1).ToString
'            '    GoTo Skip
'            'End If

'            Dim ObjReqCreateOrder As New cReqCreateOrder

'            ObjReqCreateOrder.awbId = TrackNum
'            ObjReqCreateOrder.awbIdIpp = ActualTrackNum
'            ObjReqCreateOrder.salesOrderId = OrderNo
'            ObjReqCreateOrder.serviceType = ServiceType
'            ObjReqCreateOrder.orderType = OrderType
'            ObjReqCreateOrder.channelOrder = ChannelOrder
'            ObjReqCreateOrder.vehicleType = VehicleType

'            ObjReqCreateOrder.storeCode = StoreCode
'            'If ObjReqCreateOrder.storeCode <> "" And ChannelOrder.ToUpper.Contains("INDOGROSIR") Then
'            '    ObjReqCreateOrder.storeCode = "G" & ObjReqCreateOrder.storeCode
'            'End If

'            Dim IndogrosirPakaiClusterKurWil As Boolean = False
'            If ChannelOrder.ToUpper.Contains("INDOGROSIR") Then
'                If ObjReqCreateOrder.vehicleType.ToUpper <> "MOTOR" Then
'                    'cek lokasi indogrosir, apakah main cluster kurwil indopaket
'                    Query = "select count(key1) from const where key1 = 'IGRKURWIL' and key2 = @StoreCode"
'                    Query &= " and curdate() between activedate and ifnull(inactivedate,curdate())"

'                    SqlParam = New Dictionary(Of String, String)
'                    SqlParam.Add("@StoreCode", ObjReqCreateOrder.storeCode)

'                    Try
'                        If ObjSql.ExecScalarWithParam(MCon, Query, SqlParam) > 0 Then
'                            IndogrosirPakaiClusterKurWil = True
'                        End If
'                    Catch ex As Exception
'                        IndogrosirPakaiClusterKurWil = False
'                    End Try
'                End If 'dari If ObjReqCreateOrder.vehicleType.ToUpper <> "MOTOR"
'            End If 'dari If ChannelOrder.ToUpper.Contains("INDOGROSIR")

'            If IndogrosirPakaiClusterKurWil Then
'                'perlu di-mapping ke cluster kurwil indopaket
'                Dim NewStoreCode As String = ""

'                Query = "Select Distinct w.Cluster From KurirWilayahDistrictCoverage w"
'                Query &= " Inner Join MstPostalCode p on (p.Code = @CoPostalCode And p.District = w.District)"

'                SqlParam = New Dictionary(Of String, String)
'                SqlParam.Add("@CoPostalCode", CoPostalCode)

'                Try
'                    NewStoreCode = ObjSql.ExecScalarWithParam(MCon, Query, SqlParam)
'                    If NewStoreCode = "-1" Then
'                        NewStoreCode = ""
'                    End If
'                Catch ex As Exception
'                    NewStoreCode = ""
'                End Try

'                If NewStoreCode <> "" Then
'                    ObjReqCreateOrder.storeCode = NewStoreCode
'                End If
'            End If 'dari If IndogrosirPakaiClusterKurWil

'            ObjReqCreateOrder.senderName = ActualSenderName

'            ObjReqCreateOrder.pickupPin = PINPickup
'            ObjReqCreateOrder.cancelPin = PINCancel
'            ObjReqCreateOrder.keepPin = PINKeep
'            ObjReqCreateOrder.returnPin = PINReturn

'            ObjReqCreateOrder.minDeliveryTime = ConvertTimeUtc(DeliveryMinTime)
'            ObjReqCreateOrder.maxDeliveryTime = ConvertTimeUtc(DeliveryMaxTime)

'            ObjReqCreateOrder.minDeliveryTimeOri = ConvertTimeUtc(DeliveryMinTimeOri)
'            ObjReqCreateOrder.maxDeliveryTimeOri = ConvertTimeUtc(DeliveryMaxTimeOri)


'            ObjReqCreateOrder.origin.pickupName = ShName
'            ObjReqCreateOrder.origin.pickupPhoneNumber = ShPhone

'            Dim ShAddressExt As String = ShAddress
'            If Not ShAddressExt.Contains(ShPostalCode) Then
'                If ShPostalCode.Trim <> "" Then
'                    ShAddressExt = ShAddressExt & ", " & ShPostalCode
'                End If
'            End If
'            If Not ShAddressExt.ToUpper.Contains("INDONESIA") Then
'                ShAddressExt = ShAddressExt & ", INDONESIA"
'            End If

'            ObjReqCreateOrder.origin.pickupAddress = ShAddressExt
'            ObjReqCreateOrder.origin.pickupLatitude = CDbl(ShLatitude)
'            ObjReqCreateOrder.origin.pickupLongitude = CDbl(ShLongitude)
'            ObjReqCreateOrder.origin.pickupNotes = ShNotes


'            ObjReqCreateOrder.recipient.customerName = CoName
'            ObjReqCreateOrder.recipient.customerPhoneNumber = CoPhone

'            Dim CoAddressExt As String = CoAddress
'            If Not CoAddressExt.Contains(CoPostalCode) Then
'                If CoPostalCode.Trim <> "" Then
'                    CoAddressExt = CoAddressExt & ", " & CoPostalCode
'                End If
'            End If
'            If Not CoAddressExt.ToUpper.Contains("INDONESIA") Then
'                CoAddressExt = CoAddressExt & ", INDONESIA"
'            End If

'            ObjReqCreateOrder.recipient.deliveryAddress = CoAddressExt
'            ObjReqCreateOrder.recipient.deliveryLatitude = CDbl(CoLatitude)
'            ObjReqCreateOrder.recipient.deliveryLongitude = CDbl(CoLongitude)
'            ObjReqCreateOrder.recipient.deliveryNotes = CoNotes


'            ObjReqCreateOrder.cod = FlgCOD
'            ObjReqCreateOrder.codAmount = CODValue
'            ObjReqCreateOrder.codPaymentCode = CODPaymentCode
'            ObjReqCreateOrder.codPaymentBiller = CODPaymentBiller


'            Dim IsReturnBulky As Boolean = False
'            Try
'                ObjReqCreateOrder.item = New cReqCreateOrderBulky() {}
'                ReDim ObjReqCreateOrder.item(dtReturnBulky.Rows.Count - 1)

'                For r As Integer = 0 To dtReturnBulky.Rows.Count - 1
'                    ObjReqCreateOrder.item(r) = New cReqCreateOrderBulky
'                    ObjReqCreateOrder.item(r).itemCode = dtReturnBulky.Rows(r).Item("Code").ToString.ToUpper
'                    If ObjReqCreateOrder.item(r).itemCode.Trim = "" Then
'                        ObjReqCreateOrder.item(r).itemCode = "00000000"
'                    End If
'                    ObjReqCreateOrder.item(r).itemName = dtReturnBulky.Rows(r).Item("Description").ToString.ToUpper
'                    If ObjReqCreateOrder.item(r).itemName.Trim = "" Then
'                        ObjReqCreateOrder.item(r).itemName = "BARANG KEMBALI"
'                    End If
'                    ObjReqCreateOrder.item(r).quantity = CInt(dtReturnBulky.Rows(r).Item("Qty"))

'                    If r = 0 Then
'                        IsReturnBulky = True
'                    End If
'                Next
'            Catch ex As Exception
'                ObjReqCreateOrder.item = Nothing
'                IsReturnBulky = False
'            End Try


'            ObjReqCreateOrder.bulky = False
'            Try
'                If IsBulky Or IsReturnBulky Then
'                    ObjReqCreateOrder.bulky = True
'                End If
'            Catch ex As Exception
'                ObjReqCreateOrder.bulky = False
'            End Try


'            If ObjReqCreateOrder.orderType = "PAKET-JEMPUT" Then
'                'untuk menyatakan pengirim perlu bayar ongkos kirim
'                ObjReqCreateOrder.doorPickupPaymentStatus = ShPaymentStatus
'                ObjReqCreateOrder.doorPickupPaymentValue = ShPaymentValue
'            End If


'            'ObjReqCreateOrder.recipient.deliveryDriverLocation = False
'            'Try
'            '    If CoAddInfo.Contains("PUSHNEARDST=Y") Then
'            '        ObjReqCreateOrder.recipient.deliveryDriverLocation = True
'            '    End If
'            'Catch ex As Exception
'            'End Try


'            Dim Parameter As String = JsonConvert.SerializeObject(ObjReqCreateOrder)

'            Dim AlreadyResetToken As Boolean = False
'            Dim MustResetToken As Boolean = False

'UlangGetOAuth:

'            Dim Access_Token As String = ""

'            Dim tParam() As Object
'            If MustResetToken = False Then
'                ReDim tParam(1)
'                tParam(0) = Param(0)
'                tParam(1) = Param(1)
'            Else
'                ReDim tParam(2)
'                tParam(0) = Param(0)
'                tParam(1) = Param(1)
'                tParam(2) = "1"
'                AlreadyResetToken = True
'            End If

'            Dim tResult As Object() = GetOAuth(wsAppName, wsAppVersion, tParam)
'            If tResult(0).ToString = "0" Then
'                Access_Token = tResult(2).ToString
'            Else
'                Result(1) = tResult(1).ToString
'                GoTo Skip
'            End If

'            ReDim CustomHeaders(0)
'            CustomHeaders(0) = "X-API-Key|" & Access_Token

'            Dim ParameterHeaders As String = ""
'            For h As Integer = 0 To CustomHeaders.Length - 1
'                If ParameterHeaders <> "" Then
'                    ParameterHeaders &= ","
'                End If
'                ParameterHeaders &= CustomHeaders(h)
'            Next

'            e.APIRequestLog(MCon, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "Header:" & ParameterHeaders & "-" & "Body:" & Parameter, mURLOrder, LogKeyword)

'            sDistance = "-1"
'            sPrice = "-1"
'            sOrder_Id = ""

'            Dim Response As String = ""
'            Response = ObjFungsi.SendHTTP("", "", "", mURLOrder, Parameter, "", Encoding.Default, mTimeout, "", mContentType, True, CustomHeaders)
'            Response = ("" & Response).Trim

'            Dim ResponseError As String = ""

'            Dim ObjRspOrder As New cRspCreateOrder

'            If Response <> "" Then
'                e.APIResponseLog(MCon, wsAppVersion, "WServiceDeliMan." & Method, wsUser, Response, mURLOrder, LogKeyword)
'                Try
'                    ObjRspOrder = JsonConvert.DeserializeObject(Of cRspCreateOrder)(Response)
'                    If ObjRspOrder.status = "00" Then
'                        sPrice = "1"
'                        sDistance = "1"
'                        sOrder_Id = ObjRspOrder.data.orderId

'                    ElseIf ObjRspOrder.status = "IO1001" Then
'                        'duplicate order
'                        sPrice = "1"
'                        sDistance = "1"
'                        sOrder_Id = ObjRspOrder.data.orderId

'                    Else
'                        Try
'                            ResponseError = ObjRspOrder.message
'                        Catch ex As Exception
'                        End Try
'                    End If

'                Catch ex As Exception
'                    e.ErrorLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, ex, "", LogKeyword)
'                End Try
'            Else
'                ResponseError = "No Response"
'                e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "No Response", mURLOrder, LogKeyword)
'            End If 'dari If Response <> ""

'            If sDistance = "-1" Or sPrice = "-1" Or sOrder_Id = "" Then
'                Result(1) = "Gagal CreateOrder DLM " & ResponseError
'                GoTo Skip
'            End If

'            ReDim Result(4)
'            'Result(6) = JsonConvert.SerializeObject(ObjRspPrice)
'            'Result(5) = JsonConvert.SerializeObject(ObjReqPrice)
'            Result(4) = sOrder_Id
'            Result(3) = sPrice
'            Result(2) = sDistance
'            Result(1) = ""
'            Result(0) = "0"

'            If InsertAutoOrderTracking = "1" Then
'                Query = "Insert Into AutoOrderTracking ("
'                Query &= " TrackNum, `Status`, TrackTime, TrackUser, TrackUserID, TrxPartner, ForCustomer"
'                Query &= " ) values ("
'                Query &= " '" & sOrder_Id & "', 'NEW', now(), 'DELIMAN', '" & CustomAccount3rdPartyDeliveryMan.Replace("'", "") & "', 'CREATED', 1"
'                Query &= " )"
'                ObjSql.ExecNonQuery(MCon, Query)
'            End If

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "FINISH", LogKeyword & " " & sDistance & " " & sPrice & " " & sOrder_Id)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "", ex, "", LogKeyword)

'            Result(1) = ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function

'    Public Function ConvertTimeUtc(ByVal Param As String) As String

'        Dim Result As String = ""

'        Dim ParamDateTime As Date = CDate(Param)

'        Dim TimeZone As String = ""

'        'If PostalCode <> "" Then
'        '    Dim MCon As New MySqlConnection
'        '    Try
'        '        MCon = MasterMCon.Clone
'        '        MCon.Open()

'        '        Dim Query As String = ""
'        '        Query = "Select distinct lpad(v.TimeZone,2,'0')"
'        '        Query &= " From MstPostalCode p"
'        '        Query &= " Inner Join MstProvince v on (p.Province = v.Province)"
'        '        Query &= " Where p.Code = '" & PostalCode & "'"
'        '        Query &= " Order By v.TimeZone Desc"

'        '        Dim ObjSQL As New ClsSQL
'        '        TimeZone = ("" & ObjSQL.ExecScalar(MCon, Query)).ToString.Trim

'        '    Catch ex As Exception
'        '        TimeZone = ""

'        '    Finally
'        '        If MCon.State <> ConnectionState.Closed Then
'        '            MCon.Close()
'        '        End If
'        '        Try
'        '            MCon.Dispose()
'        '        Catch ex As Exception
'        '        End Try
'        '    End Try
'        'End If

'        If TimeZone = "" Then
'            TimeZone = "07"
'        End If

'        Result = Format(ParamDateTime, "yyyy-MM-dd") & "T" & Format(ParamDateTime, "HH:mm:ss") & "+" & TimeZone & ":00"

'        Return Result

'    End Function


'    Public Class cReqEditOrder

'        Public orderId As String = "" 'Delivery ID 3rd Party

'        'Public isCOD As Boolean = False
'        Public isCOD As Object = Nothing
'        Public codAmount As Double = 0

'        Public bulky As Boolean = False
'        Public item As cReqEditOrderBulky()

'        Public returnPin As String = ""

'    End Class

'    Public Class cReqEditOrderBulky
'        Public itemCode As String = ""
'        Public itemName As String = ""
'        Public quantity As Integer = 0
'    End Class

'    Public Class cRspEditOrder
'        Public status As String = "" '00 = success
'        Public message As String = ""

'        Public data As New cRspEditOrderDetail

'        Public timestamp As String = ""
'        Public signature As String = ""
'    End Class

'    Public Class cRspEditOrderDetail
'        Public orderId As String = ""
'        Public awbId As String = ""
'        Public salesOrderId As String = ""
'        Public transactionDate As String = ""
'    End Class

'    <WebMethod()>
'    Public Function EditOrder(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object() _
'    , ByVal dsData As DataSet) As Object

'        Dim Method As String = "EditOrder"

'        Dim Result As Object() = CreateResult()
'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Dim LogKeyword As String = ""

'        Try
'            Dim Password As String = Param(1).ToString

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            Dim dtService As New DataTable
'            Dim dtCOD As New DataTable
'            Dim dtReturnBulky As New DataTable

'            For d As Integer = 0 To dsData.Tables.Count - 1
'                Select Case dsData.Tables(d).TableName.ToUpper
'                    Case "SERVICE"
'                        dtService = dsData.Tables(d).Copy

'                    Case "COD"
'                        dtCOD = dsData.Tables(d).Copy

'                    Case "RETURNBULKY"
'                        dtReturnBulky = dsData.Tables(d).Copy

'                End Select
'            Next

'            If dtService.Rows.Count < 1 Then
'                Result(1) = "Data tidak lengkap"
'                GoTo Skip
'            End If


'            Dim ObjSql As New ClsSQL

'            Dim TrackNum As String = dtService.Rows(0).Item("TrackNum").ToString
'            Dim OrderNo As String = dtService.Rows(0).Item("OrderNo").ToString
'            Dim DeliveryId As String = dtService.Rows(0).Item("DeliveryId").ToString

'            LogKeyword = OrderNo & " " & DeliveryId

'            Dim IsBulky As Boolean = False
'            Try
'                IsBulky = (dtService.Rows(0).Item("FlgBulky").ToString = "1")
'            Catch ex As Exception
'                IsBulky = False
'            End Try

'            Dim PinReturnBulky As String = ""
'            Try
'                PinReturnBulky = dtService.Rows(0).Item("PinReturnBulky").ToString
'            Catch ex As Exception
'                PinReturnBulky = ""
'            End Try

'            Dim CODValue As Double = 0
'            Try
'                CODValue = CDbl(dtCOD.Rows(0).Item("CODValue"))
'            Catch ex As Exception
'                CODValue = 0
'            End Try
'            Dim CODValueDraft As Double = -1
'            Try
'                CODValueDraft = CDbl(dtCOD.Rows(0).Item("CODValueDraft"))
'            Catch ex As Exception
'                CODValueDraft = -1
'            End Try

'            Dim ObjFungsi As New ClsFungsi


'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "START", LogKeyword)

'            Dim ObjReqEditOrder As New cReqEditOrder

'            ObjReqEditOrder.orderId = DeliveryId


'            'pindah ke bawah, setelah penentuan isCOD
'            'ObjReqEditOrder.codAmount = CODValue
'            'If Math.Round(ObjReqEditOrder.codAmount, 3) > 0 Then
'            '    ObjReqEditOrder.isCOD = True
'            'Else
'            '    ObjReqEditOrder.isCOD = False
'            'End If

'            '240327 - kasus DMS tidak bisa terima isCOD True menjadi True lagi ataupun False jadi True
'            If CODValueDraft <> -1 Then
'                '1. coddraft = 0 dan cod > 0, iscod = null, cod value = 0
'                '2. coddraft = 0 dan cod = 0, iscod = null, cod value = 0
'                '3. coddraft > 0 dan cod = 0, iscod = false, cod value = 0
'                '4. coddraft > 0 dan cod > 0, iscod = null, cod value = cod

'                '1. dan 2.
'                If Math.Round(CODValueDraft, 3) = 0 And Math.Round(CODValue, 3) >= 0 Then
'                    'ObjReqEditOrder.isCOD = Nothing
'                    CODValue = 0
'                End If

'                '3.
'                If Math.Round(CODValueDraft, 3) > 0 And Math.Round(CODValue, 3) <= 0 Then
'                    ObjReqEditOrder.isCOD = False
'                    CODValue = 0
'                End If

'                '4.
'                If Math.Round(CODValueDraft, 3) > 0 And Math.Round(CODValue, 3) > 0 Then
'                    'ObjReqEditOrder.isCOD = Nothing
'                    CODValue = CODValue
'                End If

'            End If 'dari If CODValueDraft <> -1

'            ObjReqEditOrder.codAmount = CODValue


'            ObjReqEditOrder.returnPin = PinReturnBulky

'            Dim IsReturnBulky As Boolean = False
'            Try
'                ObjReqEditOrder.item = New cReqEditOrderBulky() {}
'                ReDim ObjReqEditOrder.item(dtReturnBulky.Rows.Count - 1)

'                For r As Integer = 0 To dtReturnBulky.Rows.Count - 1
'                    ObjReqEditOrder.item(r) = New cReqEditOrderBulky
'                    ObjReqEditOrder.item(r).itemCode = dtReturnBulky.Rows(r).Item("Code").ToString.ToUpper
'                    If ObjReqEditOrder.item(r).itemCode.Trim = "" Then
'                        ObjReqEditOrder.item(r).itemCode = "00000000"
'                    End If
'                    ObjReqEditOrder.item(r).itemName = dtReturnBulky.Rows(r).Item("Description").ToString.ToUpper
'                    If ObjReqEditOrder.item(r).itemName.Trim = "" Then
'                        ObjReqEditOrder.item(r).itemName = "BARANG KEMBALI"
'                    End If
'                    ObjReqEditOrder.item(r).quantity = CInt(dtReturnBulky.Rows(r).Item("Qty"))

'                    If r = 0 Then
'                        IsReturnBulky = True
'                    End If
'                Next
'            Catch ex As Exception
'                ObjReqEditOrder.item = Nothing
'                IsReturnBulky = False
'            End Try


'            ObjReqEditOrder.bulky = False
'            Try
'                If IsBulky Or IsReturnBulky Then
'                    ObjReqEditOrder.bulky = True
'                End If
'            Catch ex As Exception
'                ObjReqEditOrder.bulky = False
'            End Try


'            Dim Parameter As String = JsonConvert.SerializeObject(ObjReqEditOrder)

'            Dim AlreadyResetToken As Boolean = False
'            Dim MustResetToken As Boolean = False

'UlangGetOAuth:

'            Dim Access_Token As String = ""

'            Dim tParam() As Object
'            If MustResetToken = False Then
'                ReDim tParam(1)
'                tParam(0) = Param(0)
'                tParam(1) = Param(1)
'            Else
'                ReDim tParam(2)
'                tParam(0) = Param(0)
'                tParam(1) = Param(1)
'                tParam(2) = "1"
'                AlreadyResetToken = True
'            End If

'            Dim tResult As Object() = GetOAuth(wsAppName, wsAppVersion, tParam)
'            If tResult(0).ToString = "0" Then
'                Access_Token = tResult(2).ToString
'            Else
'                Result(1) = tResult(1).ToString
'                GoTo Skip
'            End If

'            ReDim CustomHeaders(0)
'            CustomHeaders(0) = "X-API-Key|" & Access_Token

'            Dim ParameterHeaders As String = ""
'            For h As Integer = 0 To CustomHeaders.Length - 1
'                If ParameterHeaders <> "" Then
'                    ParameterHeaders &= ","
'                End If
'                ParameterHeaders &= CustomHeaders(h)
'            Next

'            e.APIRequestLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "Header:" & ParameterHeaders & "-" & "Body:" & Parameter, mURLEditOrder, LogKeyword)

'            Dim Response As String = ""
'            Response = ObjFungsi.SendHTTP("", "", "", mURLEditOrder, Parameter, "", Encoding.Default, "60", "", mContentType, True, CustomHeaders, "PUT")
'            Response = ("" & Response).Trim

'            Dim ResponseError As String = ""

'            Dim ObjRspOrder As New cRspEditOrder

'            If Response <> "" Then
'                e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, Response, mURLEditOrder, LogKeyword)
'                Try
'                    ObjRspOrder = JsonConvert.DeserializeObject(Of cRspEditOrder)(Response)
'                    If ObjRspOrder.status = "00" Then

'                    Else
'                        Try
'                            ResponseError = ObjRspOrder.message
'                        Catch ex As Exception
'                        End Try
'                    End If

'                Catch ex As Exception
'                    e.ErrorLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, ex, "", LogKeyword)
'                End Try
'            Else
'                ResponseError = "No Response"
'                e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "No Response", mURLOrder, LogKeyword)
'            End If 'dari If Response <> ""

'            If ResponseError <> "" Then
'                Result(1) = "Gagal EditOrder DLM " & ResponseError
'                GoTo Skip
'            End If

'            ReDim Result(2)
'            Result(2) = ""
'            Result(1) = "OK"
'            Result(0) = "0"

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "FINISH", LogKeyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "", ex, "", LogKeyword)

'            Result(1) = ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Public Class cCallbackTrackingReq
'        Public data As cCallbackTrackingReqDetail
'        Public timestamp As String = ""
'        Public signature As String = ""
'    End Class

'    Public Class cCallbackTrackingReqDetail

'        Public orderId As String = "" 'delivery third party
'        Public awbId As String = "" 'awb indopaket
'        Public salesOrderId As String = "" 'order partner

'        Public status As String = ""
'        Public trackDate As String = ""

'        Public delimanId As String = "" 'driver
'        Public delimanName As String = ""
'        Public delimanPhone As String = ""
'        Public delimanVehicle As String = ""
'        Public delimanCompany As String = ""

'        Public receiverName As String = "" 'penerima
'        Public receiverRelationship As String = ""

'        Public deliveryPow As String = "" 'bukti foto
'        Public codPow As String = ""
'        Public bulkyPow As String = ""

'        Public cancelCode As String = "" 'pengiriman dibatalkan
'        Public cancelReason As String = ""

'        Public failReason As String = "" 'pengiriman gagal

'        Public orderTrackingUrl As String = ""

'    End Class

'    <WebMethod()>
'    Public Function CallbackTracking(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "CallbackTracking"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString

'            Dim Request As String = Param(2).ToString
'            Request = Request.Trim

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If Request = "" Then
'                Result(1) = "Input request"
'                GoTo Skip
'            End If

'            Dim Keyword As String = Format(Date.Now, "yyMMddHHmmss")

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "MULAI " & Request, Keyword)


'            Dim ObjRequest As cCallbackTrackingReq = JsonConvert.DeserializeObject(Of cCallbackTrackingReq)(Request)

'            If ObjRequest Is Nothing Then
'                Result(1) = "Gagal convert parameter request"
'                GoTo Skip
'            End If


'            Dim ObjSQL As New ClsSQL

'            Dim ThirdPartyList As String = CustomAccount3rdPartyDeliveryMan
'            If CustomAccount3rdPartyIndopaketMotor <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketMotor
'            End If
'            If CustomAccount3rdPartyIndopaketMobil <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketMobil
'            End If
'            If CustomAccount3rdPartyIndopaketInstantMotor <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketInstantMotor
'            End If
'            If CustomAccount3rdPartyIndopaketInstantMobil <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketInstantMobil
'            End If

'            'validasi data
'            Query = "Select count(TrackNum)"
'            Query &= " From AutoOrderTrackingHistory"
'            Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'            Query &= " And TrackUserId in (" & ThirdPartyList & ")"

'            'perlu ada pemisahan antara kiriman instant (main) dengan kiriman last mile - update 240424
'            If ObjSQL.ExecScalar(MCon, Query) > 0 Then
'                CallbackTrackingMain(AppName, AppVersion, Param)
'            Else
'                Query = "Select count(TrackNum)"
'                Query &= " From ExpAutoOrderTrackingHistory"
'                Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'                Query &= " And TrackUserId in (" & CustomAccountExpeditionDMSMotor & "," & CustomAccountExpeditionDMSMobil & ")"

'                If ObjSQL.ExecScalar(MCon, Query) > 0 Then
'                    CallbackTrackingLastMile(AppName, AppVersion, Param)
'                Else
'                    Result(1) = "Nomor " & ObjRequest.data.orderId & " tidak ditemukan"
'                    GoTo Skip
'                End If

'            End If

'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = ""

'            Process = True

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "SELESAI " & Process.ToString, Keyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    <WebMethod()>
'    Public Function CallbackTrackingMain(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "CallbackTrackingMain"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString

'            Dim Request As String = Param(2).ToString
'            Request = Request.Trim

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If Request = "" Then
'                Result(1) = "Input request"
'                GoTo Skip
'            End If

'            Dim Keyword As String = Format(Date.Now, "yyMMddHHmmss")

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "MULAI " & Request, Keyword)


'            Dim ObjRequest As cCallbackTrackingReq = JsonConvert.DeserializeObject(Of cCallbackTrackingReq)(Request)

'            If ObjRequest Is Nothing Then
'                Result(1) = "Gagal convert parameter request"
'                GoTo Skip
'            End If


'            Dim ObjSQL As New ClsSQL

'            Dim ThirdPartyList As String = CustomAccount3rdPartyDeliveryMan
'            If CustomAccount3rdPartyIndopaketMotor <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketMotor
'            End If
'            If CustomAccount3rdPartyIndopaketMobil <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketMobil
'            End If
'            If CustomAccount3rdPartyIndopaketInstantMotor <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketInstantMotor
'            End If
'            If CustomAccount3rdPartyIndopaketInstantMobil <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketInstantMobil
'            End If

'            'validasi data
'            Query = "Select count(TrackNum)"
'            Query &= " From AutoOrderTrackingHistory"
'            Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'            Query &= " And TrackUserId in (" & ThirdPartyList & ")"

'            If ObjSQL.ExecScalar(MCon, Query) < 1 Then
'                Result(1) = "Nomor " & ObjRequest.data.orderId & " tidak ditemukan"
'                GoTo Skip
'            End If


'            Dim StatusTime As String = ""
'            Try
'                StatusTime = Format(CDate(ObjRequest.data.trackDate), "yyyy-MM-dd HH:mm:ss")
'            Catch ex As Exception
'                StatusTime = ObjSQL.ExecScalar(MCon, "select cast(now() as char)").ToString
'            End Try

'            Dim CancelCode As String = ""
'            Try
'                CancelCode = ("" & ObjRequest.data.cancelCode)
'            Catch ex As Exception
'                CancelCode = ""
'            End Try
'            Dim CancelType As String = CallbackTrackingCancelType(ObjRequest.data.status, CancelCode)

'            Dim PrevStatus As String = ""
'            Dim StatusIPP As String = CallbackTrackingMappingStatus(IIf(CancelType <> "", CancelType, ObjRequest.data.status), PrevStatus)
'            If StatusIPP = "" Then
'                Result(1) = "Gagal mapping status " & ObjRequest.data.status & " " & ObjRequest.data.orderId
'                GoTo Skip
'            End If

'            ObjRequest.data.delimanName = ("" & ObjRequest.data.delimanName)
'            ObjRequest.data.delimanName = WService.StrSQLRemoveChar(ObjRequest.data.delimanName)
'            If StatusIPP = "HDV" Then
'                If ObjRequest.data.delimanName = "" Then
'                    Result(1) = "Input informasi kurir " & ObjRequest.data.orderId
'                    GoTo Skip
'                End If

'                If PrevStatus <> "" Then

'                    Query = "Select count(TrackNum)"
'                    Query &= " From AutoOrderCallbackTracking"
'                    Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'                    Query &= " And ThirdParty in (" & ThirdPartyList & ")"
'                    Query &= " And TrackStatus = '" & PrevStatus & "'"
'                    If ObjSQL.ExecScalar(MCon, Query) < 1 Then
'                        Result(1) = "Belum ada status " & PrevStatus & " " & ObjRequest.data.orderId
'                        GoTo Skip
'                    End If

'                End If

'            End If

'            ObjRequest.data.receiverName = ("" & ObjRequest.data.receiverName)
'            ObjRequest.data.receiverName = WService.StrSQLRemoveChar(ObjRequest.data.receiverName)
'            If StatusIPP = "HDS" Then
'                If ObjRequest.data.receiverName = "" Then
'                    Result(1) = "Input informasi penerima " & ObjRequest.data.orderId
'                    GoTo Skip
'                End If

'                If PrevStatus <> "" Then

'                    Query = "Select count(TrackNum)"
'                    Query &= " From AutoOrderCallbackTracking"
'                    Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'                    Query &= " And ThirdParty in (" & ThirdPartyList & ")"
'                    Query &= " And TrackStatus = '" & PrevStatus & "'"
'                    If ObjSQL.ExecScalar(MCon, Query) < 1 Then
'                        Result(1) = "Belum ada status " & PrevStatus & " " & ObjRequest.data.orderId
'                        GoTo Skip
'                    End If

'                End If

'            End If


'            If StatusIPP = "CRI" Or StatusIPP = "CPY" Then

'                If PrevStatus <> "" Then

'                    Query = "Select count(TrackNum)"
'                    Query &= " From AutoOrderCallbackTracking"
'                    Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'                    Query &= " And ThirdParty in (" & ThirdPartyList & ")"
'                    Query &= " And TrackStatus = '" & PrevStatus & "'"
'                    If ObjSQL.ExecScalar(MCon, Query) < 1 Then
'                        Result(1) = "Belum ada status " & PrevStatus & " " & ObjRequest.data.orderId
'                        GoTo Skip
'                    End If

'                End If

'            End If

'            Query = "Select TrackUserId"
'            Query &= " From AutoOrderTrackingHistory"
'            Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'            Query &= " And TrackUserId in (" & ThirdPartyList & ")"
'            Dim TrackUserId As String = ObjSQL.ExecScalar(MCon, Query).ToString


'            Request = JsonConvert.SerializeObject(ObjRequest)


'            'insert ke table antrian
'            Query = "Insert into AutoOrderCallbackTracking ("
'            Query &= " ThirdParty, Parameter"
'            Query &= " , TrackNum, TrackStatus"
'            Query &= " , AddTime, AddUser"
'            Query &= " ) values ("
'            Query &= " '" & TrackUserId & "', '" & Request.Replace("'", "") & "'"
'            Query &= " , '" & ObjRequest.data.orderId & "', '" & ObjRequest.data.status & "'"
'            Query &= " , '" & StatusTime & "', 'NEW'"
'            Query &= " )"
'            If ObjSQL.ExecNonQuery(MCon, Query) = False Then
'                Result(1) = "Gagal insert antrian " & ObjRequest.data.orderId
'                GoTo Skip
'            End If

'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = ""

'            Process = True

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "SELESAI " & Process.ToString, Keyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function

'    <WebMethod()>
'    Public Sub CallbackTrackingProcess(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object())

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "CallbackTrackingProcess"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            Dim Keyword As String = Format(Date.Now, "yyMMddHHmmss")

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "MULAI", Keyword)

'            Dim ObjSQL As New ClsSQL

'            Dim ThirdPartyList As String = CustomAccount3rdPartyDeliveryMan
'            If CustomAccount3rdPartyIndopaketMotor <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketMotor
'            End If
'            If CustomAccount3rdPartyIndopaketMobil <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketMobil
'            End If
'            If CustomAccount3rdPartyIndopaketInstantMotor <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketInstantMotor
'            End If
'            If CustomAccount3rdPartyIndopaketInstantMobil <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketInstantMobil
'            End If

'            Query = "Select TrackId, cast(Parameter as char) as Request"
'            Query &= " , cast(AddTime as char) as TrackTime"
'            Query &= " From AutoOrderCallbackTracking"
'            Query &= " Where `Status` = '0'"
'            Query &= " And ThirdParty in (" & ThirdPartyList & ")"
'            Query &= " Order By AddTime, Retry"
'            Dim dtProcess As DataTable = ObjSQL.ExecDatatable(MCon, Query)

'            For p As Integer = 0 To dtProcess.Rows.Count - 1

'                Dim TrackId As String = dtProcess.Rows(p).Item("TrackId").ToString

'                Try
'                    Dim Request As String = dtProcess.Rows(p).Item("Request").ToString

'                    Dim ErrorMessage As String = ""

'                    Dim ObjRequest As cCallbackTrackingReq = JsonConvert.DeserializeObject(Of cCallbackTrackingReq)(Request)

'                    If ObjRequest Is Nothing Then
'                        ErrorMessage = "Gagal convert parameter request " & TrackId
'                        GoTo SkipNext
'                    End If

'                    Dim CancelCode As String = ""
'                    Try
'                        CancelCode = ("" & ObjRequest.data.cancelCode)
'                    Catch ex As Exception
'                        CancelCode = ""
'                    End Try
'                    Dim CancelType As String = CallbackTrackingCancelType(ObjRequest.data.status, CancelCode)

'                    'validasi
'                    Dim StatusIPP As String = CallbackTrackingMappingStatus(IIf(CancelType <> "", CancelType, ObjRequest.data.status), "")
'                    If StatusIPP = "" Then
'                        ErrorMessage = "Gagal mapping status " & ObjRequest.data.status & " " & TrackId
'                        GoTo SkipNext
'                    End If

'                    If StatusIPP = "CRI" Then
'                        Try
'                            If ("" & ObjRequest.data.codPow).ToString.Trim <> "" Then
'                                StatusIPP = "CPY"
'                            End If
'                        Catch ex As Exception
'                        End Try
'                    End If

'                    Query = "Select count(TrackNum)"
'                    Query &= " From AutoOrderTrackingHistory"
'                    Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'                    Query &= " And TrackUserId in (" & ThirdPartyList & ")"
'                    If ObjSQL.ExecScalar(MCon, Query) < 1 Then
'                        ErrorMessage = "Nomor " & ObjRequest.data.orderId & " tidak ditemukan " & TrackId
'                        GoTo SkipNext
'                    End If

'                    Dim TrackTime As String = ""
'                    Try
'                        TrackTime = Format(CDate(ObjRequest.data.trackDate), "yyyy-MM-dd HH:mm:ss")
'                    Catch ex As Exception
'                        TrackTime = dtProcess.Rows(p).Item("TrackTime").ToString
'                    End Try

'                    Query = "Select count(TrackNum)"
'                    Query &= " From AutoOrderTracking"
'                    Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'                    Query &= " And TrackUserId in (" & ThirdPartyList & ")"
'                    Query &= " And `Status` = '" & StatusIPP & "'"
'                    'Query &= " And TrackTime = '" & TrackTime & "'" '-- tidak digunakan per 231228
'                    If ObjSQL.ExecScalar(MCon, Query) > 0 Then
'                        ErrorMessage = "" 'sudah pernah dilakukan
'                        GoTo SkipNext
'                    End If

'                    Query = "Select TrackUserId"
'                    Query &= " From AutoOrderTrackingHistory"
'                    Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'                    Query &= " And TrackUserId in (" & ThirdPartyList & ")"
'                    Dim TrackUserId As String = ObjSQL.ExecScalar(MCon, Query).ToString


'                    ObjRequest.data.orderTrackingUrl = ("" & ObjRequest.data.orderTrackingUrl)

'                    ObjRequest.data.delimanId = ("" & ObjRequest.data.delimanId)
'                    ObjRequest.data.delimanId = Left(ObjRequest.data.delimanId, 45)

'                    ObjRequest.data.delimanName = ("" & ObjRequest.data.delimanName)
'                    ObjRequest.data.delimanName = WService.StrSQLRemoveChar(Left(ObjRequest.data.delimanName, 100))

'                    ObjRequest.data.delimanVehicle = ("" & ObjRequest.data.delimanVehicle)
'                    ObjRequest.data.delimanVehicle = Left(ObjRequest.data.delimanVehicle, 45)

'                    Dim DriverCompany As String = ""
'                    Try
'                        DriverCompany = ObjRequest.data.delimanCompany.Trim.ToUpper
'                    Catch ex As Exception
'                        DriverCompany = ""
'                    End Try

'                    ObjRequest.data.receiverName = ("" & ObjRequest.data.receiverName)
'                    ObjRequest.data.receiverName = WService.StrSQLRemoveChar(Left(ObjRequest.data.receiverName, 100))

'                    ObjRequest.data.receiverRelationship = ("" & ObjRequest.data.receiverRelationship)
'                    ObjRequest.data.receiverRelationship = Left(ObjRequest.data.receiverRelationship, 45)


'                    Dim AlreadyOOCFLR As Boolean = False
'                    Query = "Select count(TrackNum) From AutoOrderTracking"
'                    Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'                    Query &= " And TrackUserId in (" & ThirdPartyList & ")"
'                    Query &= " And `Status` in ('OOC','FLR')"
'                    If ObjSQL.ExecScalar(MCon, Query) > 0 Then
'                        AlreadyOOCFLR = True 'sudah pernah ada status OOC / FLR, maka status lain tidak berlaku
'                    End If


'                    'insert data
'                    Dim InsertAutoOrderTracking As Boolean = True

'                    If StatusIPP = "CRI" Or StatusIPP = "CPY" Then
'                        InsertAutoOrderTracking = False 'tidak perlu insert ke AutoOrderTracking
'                    End If 'dari If StatusIPP = "CRI" Or StatusIPP = "CPY"

'                    If AlreadyOOCFLR = True Then
'                        InsertAutoOrderTracking = False
'                    End If

'                    If StatusIPP = "NRD" Then
'                        InsertAutoOrderTracking = False
'                    End If

'                    If InsertAutoOrderTracking Then

'                        Query = "Insert Into AutoOrderTracking ("
'                        Query &= " TrackNum, `Status`, TrxPartner"
'                        Query &= " , TrackTime, TrackUser, TrackUserID"
'                        Query &= " , ForCustomer"

'                        If StatusIPP = "CRI" Or StatusIPP = "CPY" Then
'                            Query &= " , OprID1, OprName1" '-- nama yang melakukan proses pengembalian bulky atau setor COD
'                        End If
'                        If (StatusIPP = "OOC" Or StatusIPP = "FLR") And ObjRequest.data.delimanName <> "" Then
'                            Query &= " , OprID1, OprName1" '-- nama pengirim terakhir, sebelum gagal kirim / gagal proses
'                        End If
'                        If StatusIPP = "PUL" Then
'                            Query &= " , OprID1, OprName1" '-- informasi driver
'                        End If
'                        If StatusIPP = "HDV" Or StatusIPP = "HDS" Then
'                            Query &= " , OprID1, OprName1" '-- nama pengirim
'                        End If
'                        If StatusIPP = "HDS" Then
'                            Query &= " , OprID2, OprName2" '-- nama penerima
'                        End If

'                        If StatusIPP = "HDS" Or StatusIPP = "CRI" Or StatusIPP = "CPY" Or StatusIPP = "OOC" Or StatusIPP = "FLR" Then
'                            Query &= " , AddInfo" '-- foto/bukti/alasan proses
'                        End If

'                        Query &= " ) values ("
'                        Query &= " '" & ObjRequest.data.orderId & "', '" & StatusIPP & "', '" & ObjRequest.data.status.ToUpper & "'"
'                        Query &= " , '" & TrackTime & "', 'DELIMAN', '" & TrackUserId & "'"

'                        If StatusIPP = "PUQ" Or StatusIPP = "PUL" Or StatusIPP = "HDV" Or StatusIPP = "HDS" Then
'                            Query &= " , '1'" 'ForCustomer
'                        Else
'                            Query &= " , '0'"
'                        End If

'                        If StatusIPP = "CRI" Or StatusIPP = "CPY" Then
'                            Query &= " , '" & ObjRequest.data.delimanId & "', '" & ObjRequest.data.delimanName & "'" '-- nama yang melakukan proses pengembalian bulky atau setor COD
'                        End If
'                        If (StatusIPP = "OOC" Or StatusIPP = "FLR") And ObjRequest.data.delimanName <> "" Then
'                            Query &= " , '" & ObjRequest.data.delimanId & "', '" & ObjRequest.data.delimanName & "'" '-- nama pengirim terakhir, sebelum gagal kirim / gagal proses
'                        End If
'                        If StatusIPP = "PUL" Then
'                            Query &= " , '" & ("" & ObjRequest.data.delimanPhone) & "', '" & ObjRequest.data.delimanName & "'" '-- nama pengirim
'                        End If
'                        If StatusIPP = "HDV" Or StatusIPP = "HDS" Then
'                            Query &= " , '" & ObjRequest.data.delimanId & "', '" & ObjRequest.data.delimanName & "'" '-- nama pengirim
'                        End If
'                        If StatusIPP = "HDS" Then
'                            Query &= " , '" & ObjRequest.data.receiverRelationship & "', '" & ObjRequest.data.receiverName & "'" '-- nama penerima
'                        End If


'                        Dim Reason As String = ""
'                        Try
'                            Reason = ("" & ObjRequest.data.failReason).ToString
'                        Catch ex As Exception
'                        End Try
'                        Try
'                            Reason = ("" & ObjRequest.data.cancelCode).ToString & "|" & ("" & ObjRequest.data.cancelReason).ToString
'                        Catch ex As Exception
'                        End Try

'                        If StatusIPP = "HDS" Then
'                            Query &= " , '" & ("" & ObjRequest.data.deliveryPow).ToString & "'" '-- bukti diterima
'                        ElseIf StatusIPP = "CRI" Then
'                            Query &= " , '" & ("" & ObjRequest.data.bulkyPow).ToString & "'" '-- bukti pengembalian barang
'                        ElseIf StatusIPP = "CPY" Then
'                            Query &= " , '" & ("" & ObjRequest.data.codPow).ToString & "'" '-- bukti proses setor pembayaran
'                        ElseIf StatusIPP = "OOC" Then
'                            Query &= " , '" & Reason & "'" '-- alasan gagal kirim
'                        ElseIf StatusIPP = "FLR" Then
'                            Query &= " , '" & Reason & "'" '-- alasan gagal proses
'                        End If

'                        Query &= " )"
'                        If ObjSQL.ExecNonQuery(MCon, Query) = False Then
'                            ErrorMessage = "Gagal insert autoordertracking " & TrackId
'                            GoTo Skipnext
'                        End If

'                    End If 'dari If InsertAutoOrderTracking


'                    If StatusIPP = "CRI" Or StatusIPP = "CPY" Or StatusIPP = "OOC" Then

'                        'insert ke table PushTracking
'                        Dim InsertPushTracking As Boolean = False

'                        Dim OrderNo As String = ObjRequest.data.salesOrderId.Trim.ToUpper

'                        Query = "Select Account, TrackNum"
'                        Query &= " From AutoOrderThirdPartyTransaction"
'                        Query &= " Where dTrackNum = '" & ObjRequest.data.awbId.Trim.ToUpper & "' And TrackNum <> ''"

'                        Dim dtAccountTrackNum As DataTable = ObjSQL.ExecDatatable(MCon, Query)
'                        Dim AccountRep As String = ""
'                        Dim TrackNumRep As String = ""

'                        Try
'                            AccountRep = dtAccountTrackNum.Rows(0).Item("Account").ToString
'                            TrackNumRep = dtAccountTrackNum.Rows(0).Item("TrackNum").ToString
'                        Catch ex As Exception
'                            AccountRep = ""
'                            TrackNumRep = ""
'                        End Try

'                        If AccountRep = "" Then
'                            AccountRep = "-------"
'                        End If

'                        Dim SkipPushTracking As Boolean = False

'                        If CustomAccountKlikIStore <> "" And CustomAccountKlikIStore.Contains(AccountRep) Then
'                            InsertPushTracking = True
'                        End If
'                        If CustomAccountKlikFood <> "" And CustomAccountKlikFood.Contains(AccountRep) Then
'                            InsertPushTracking = True
'                        End If
'                        If CustomAccountKlikApka <> "" And CustomAccountKlikApka.Contains(AccountRep) Then
'                            InsertPushTracking = True
'                            SkipPushTracking = True
'                        End If
'                        If CustomAccountKlikReOrder <> "" And CustomAccountKlikReOrder.Contains(AccountRep) Then
'                            InsertPushTracking = True
'                        End If

'                        If InsertPushTracking Then

'                            'perlu dibedakan dari yang perlu di-hit ke Klik
'                            Dim AccountRepSvrInventory As String = CustomAccountServerInventoryDelimanCode.Replace("'", "") & AccountRep

'                            Try
'                                If ("" & ObjRequest.data.bulkyPow).ToString.Trim <> "" Then
'                                    'Insert Push Tracking DeliveryMan CRI
'                                    Query = "Insert Into PushTracking ("
'                                    Query &= " `Account`, TrackNum, OrderNo, TrackStatus"
'                                    Query &= " , TrackTime, TrackUser, TrackUserID"
'                                    Query &= " , Company1, Opr1, OprID1, OprName1"
'                                    'Query &= " , Company2, Opr2, OprID2, OprName2"
'                                    'Query &= " , Deskripsi, Description"
'                                    Query &= " , PushStatus, AddTime, AddUser"
'                                    Query &= " ) Select"
'                                    Query &= " '" & AccountRepSvrInventory & "', '" & TrackNumRep & "', '" & OrderNo & "', '" & "CRI" & "'"
'                                    Query &= " , '" & TrackTime & "', '" & "DLM" & "', '" & "DLM" & "'"
'                                    Query &= " , '" & "DLM" & "', '" & "DLM" & "', '" & ObjRequest.data.delimanId & "', '" & ObjRequest.data.delimanName & "'"
'                                    'Query &= " , '" & "" & "', '" & "" & "', '" & "" & "', '" & "" & "'"
'                                    'Query &= " , '" & ForCustomerDesc(0) & "', '" & ForCustomerDesc(1) & "'"
'                                    Query &= " , '0', now(), '" & User & "'"
'                                    ObjSQL.ExecNonQuery(MCon, Query)
'                                End If
'                            Catch ex As Exception
'                                e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, "Insert PushTracking CPY", TrackId)
'                            End Try


'                            Try
'                                If ("" & ObjRequest.data.codPow).ToString.Trim <> "" Then
'                                    'Insert Push Tracking DeliveryMan CPY
'                                    Query = "Insert Into PushTracking ("
'                                    Query &= " `Account`, TrackNum, OrderNo, TrackStatus"
'                                    Query &= " , TrackTime, TrackUser, TrackUserID"
'                                    Query &= " , Company1, Opr1, OprID1, OprName1"
'                                    'Query &= " , Company2, Opr2, OprID2, OprName2"
'                                    'Query &= " , Deskripsi, Description"
'                                    Query &= " , PushStatus, AddTime, AddUser"
'                                    Query &= " ) Select"
'                                    Query &= " '" & AccountRepSvrInventory & "', '" & TrackNumRep & "', '" & OrderNo & "', '" & "CPY" & "'"
'                                    Query &= " , '" & TrackTime & "', '" & "DLM" & "', '" & "DLM" & "'"
'                                    Query &= " , '" & "DLM" & "', '" & "DLM" & "', '" & ObjRequest.data.delimanId & "', '" & ObjRequest.data.delimanName & "'"
'                                    'Query &= " , '" & "" & "', '" & "" & "', '" & "" & "', '" & "" & "'"
'                                    'Query &= " , '" & ForCustomerDesc(0) & "', '" & ForCustomerDesc(1) & "'"
'                                    Query &= " , '0', now(), '" & User & "'"
'                                    ObjSQL.ExecNonQuery(MCon, Query)
'                                End If
'                            Catch ex As Exception
'                                e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, "Insert PushTracking CPY", TrackId)
'                            End Try


'                            Try
'                                If ("" & ObjRequest.data.cancelCode).ToString.Trim <> "" Then

'                                    Dim CancelReason As String = ("" & ObjRequest.data.cancelCode).ToString.Trim.ToUpper
'                                    Try
'                                        CancelReason &= "|" & ("" & ObjRequest.data.cancelReason).ToString.Trim.ToUpper
'                                    Catch ex As Exception
'                                    End Try

'                                    'Insert Push Tracking DeliveryMan OOC, ke Server Inventory
'                                    Query = "Insert Into PushTracking ("
'                                    Query &= " `Account`, TrackNum, OrderNo, TrackStatus"
'                                    Query &= " , TrackTime, TrackUser, TrackUserID"
'                                    Query &= " , Company1, Opr1, OprID1, OprName1"
'                                    'Query &= " , Company2, Opr2, OprID2, OprName2"
'                                    'Query &= " , Deskripsi, Description"
'                                    Query &= " , AddInfo"
'                                    Query &= " , PushStatus, AddTime, AddUser"
'                                    Query &= " ) Select"
'                                    Query &= " '" & AccountRepSvrInventory & "', '" & TrackNumRep & "', '" & OrderNo & "', '" & "OOC" & "'"
'                                    Query &= " , '" & TrackTime & "', '" & "DLM" & "', '" & "DLM" & "'"
'                                    Query &= " , '" & "DLM" & "', '" & "DLM" & "', '" & ObjRequest.data.delimanId & "', '" & ObjRequest.data.delimanName & "'"
'                                    'Query &= " , '" & "" & "', '" & "" & "', '" & "" & "', '" & "" & "'"
'                                    'Query &= " , '" & ForCustomerDesc(0) & "', '" & ForCustomerDesc(1) & "'"
'                                    Query &= " , '" & CancelReason & "'"
'                                    Query &= " , '0', now(), '" & User & "'"
'                                    ObjSQL.ExecNonQuery(MCon, Query)


'                                    If SkipPushTracking = False Then
'                                        'Insert Push Tracking DeliveryMan OOC, ke KlikIndomaret
'                                        Query = "Insert Into PushTracking ("
'                                        Query &= " `Account`, TrackNum, OrderNo, TrackStatus"
'                                        Query &= " , TrackTime, TrackUser, TrackUserID"
'                                        Query &= " , Company1, Opr1, OprID1, OprName1"
'                                        'Query &= " , Company2, Opr2, OprID2, OprName2"
'                                        'Query &= " , Deskripsi, Description"
'                                        Query &= " , AddInfo"
'                                        Query &= " , PushStatus, AddTime, AddUser"
'                                        Query &= " ) Select"
'                                        Query &= " '" & AccountRep & "', '" & TrackNumRep & "', '" & OrderNo & "', '" & "HDX" & "'"
'                                        Query &= " , '" & TrackTime & "', '" & "DLM" & "', '" & "DLM" & "'"
'                                        Query &= " , '" & "DLM" & "', '" & "DLM" & "', '" & ObjRequest.data.delimanId & "', '" & ObjRequest.data.delimanName & "'"
'                                        'Query &= " , '" & "" & "', '" & "" & "', '" & "" & "', '" & "" & "'"
'                                        'Query &= " , '" & ForCustomerDesc(0) & "', '" & ForCustomerDesc(1) & "'"
'                                        Query &= " , '" & CancelReason & "'"
'                                        Query &= " , '0', now(), '" & User & "'"
'                                        ObjSQL.ExecNonQuery(MCon, Query)
'                                    End If

'                                End If
'                            Catch ex As Exception
'                                e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, "Insert PushTracking CPY", TrackId)
'                            End Try

'                        End If 'dari If InsertPushTracking

'                    End If 'dari If StatusIPP = "CRI" Or StatusIPP = "CPY" Or StatusIPP = "OOC"


'                    If StatusIPP = "NRD" Then
'                        'Near Destination
'                    End If 'dari If StatusIPP = "NRD"


'                    If AlreadyOOCFLR = False And StatusIPP = "PUQ" Then

'                        Try
'                            Query = "Select TrackNum From AutoOrderThirdPartyTransaction"
'                            Query &= " Where ThirdParty = '" & TrackUserId & "' And dOrderNo = '" & ObjRequest.data.orderId & "'"
'                            Dim TrackNumIPP As String = ObjSQL.ExecScalar(MCon, Query)
'                            If TrackNumIPP <> "" Then

'                                '2 = TrackNumIpp, 3 = ThirdParty, 4 = DriverId, 5 = DriverName, 6 = DriverPhone, 7 = VehicleNo, 8 = TrackingUrl
'                                Dim xParam(8) As Object
'                                xParam(2) = TrackNumIPP
'                                xParam(3) = TrackUserId
'                                xParam(4) = ObjRequest.data.delimanId
'                                xParam(5) = ObjRequest.data.delimanName
'                                xParam(6) = ObjRequest.data.delimanPhone
'                                xParam(7) = ObjRequest.data.delimanVehicle
'                                xParam(8) = ObjRequest.data.orderTrackingUrl

'                                WService.LiveTrackingUpdate(MCon, xParam)

'                            End If 'dari If TrackNumIPP <> ""
'                        Catch ex As Exception
'                        End Try

'                    End If


'                    If AlreadyOOCFLR = False And StatusIPP = "PUL" Then

'                        Try
'                            Query = "Select TrackNum From AutoOrderThirdPartyTransaction"
'                            Query &= " Where ThirdParty = '" & TrackUserId & "' And dOrderNo = '" & ObjRequest.data.orderId & "'"
'                            Dim TrackNumIPP As String = ObjSQL.ExecScalar(MCon, Query)
'                            If TrackNumIPP <> "" Then

'                                '2 = TrackNumIpp, 3 = ThirdParty, 4 = DriverId, 5 = DriverName, 6 = DriverPhone, 7 = VehicleNo, 8 = TrackingUrl
'                                Dim xParam(8) As Object
'                                xParam(2) = TrackNumIPP
'                                xParam(3) = TrackUserId
'                                xParam(4) = ObjRequest.data.delimanId
'                                xParam(5) = ObjRequest.data.delimanName
'                                xParam(6) = ObjRequest.data.delimanPhone
'                                xParam(7) = ObjRequest.data.delimanVehicle
'                                xParam(8) = ObjRequest.data.orderTrackingUrl

'                                WService.LiveTrackingUpdate(MCon, xParam)

'                            End If 'dari If TrackNumIPP <> ""
'                        Catch ex As Exception
'                        End Try

'                    End If

'                    If AlreadyOOCFLR = False And StatusIPP = "HDV" Then
'                        Query = "Update AutoOrderThirdPartyTransaction"
'                        Query &= " Set DriverName = '" & ObjRequest.data.delimanName & "'"
'                        Query &= " , DriverPhone = '" & ObjRequest.data.delimanPhone & "'"
'                        Query &= " , DriverId = '" & ObjRequest.data.delimanId & "'"
'                        Query &= " , DriverVehicleNo = '" & ObjRequest.data.delimanVehicle & "'"
'                        Query &= " , DriverCompany = '" & DriverCompany & "'"
'                        Query &= " , UpdTime2 = now(), UpdUser2 = 'UPD-HDV'"
'                        Query &= " Where ThirdParty = '" & TrackUserId & "'"
'                        Query &= " And dOrderNo = '" & ObjRequest.data.orderId & "'"
'                        Query &= " And DriverName = ''"
'                        ObjSQL.ExecNonQuery(MCon, Query)

'                        Try
'                            Query = "Select Account, TrackNum From AutoOrderThirdPartyTransaction"
'                            Query &= " Where ThirdParty = @TrackUserId And dOrderNo = @orderId"

'                            SqlParam = New Dictionary(Of String, String)
'                            SqlParam.Add("@TrackUserId", TrackUserId)
'                            SqlParam.Add("@orderId", ObjRequest.data.orderId)

'                            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)
'                            Dim AccountIPP As String = ""
'                            Dim TrackNumIPP As String = ""
'                            If Not IsNothing(dtQuery) Then
'                                If dtQuery.Rows.Count > 0 Then
'                                    AccountIPP = dtQuery.Rows(0).Item("Account").ToString
'                                    TrackNumIPP = dtQuery.Rows(0).Item("TrackNum").ToString
'                                End If
'                            End If

'                            If TrackNumIPP <> "" Then

'                                '2 = TrackNumIpp, 3 = ThirdParty, 4 = DriverId, 5 = DriverName, 6 = DriverPhone, 7 = VehicleNo, 8 = TrackingUrl
'                                Dim xParam(8) As Object
'                                xParam(2) = TrackNumIPP
'                                xParam(3) = TrackUserId
'                                xParam(4) = ObjRequest.data.delimanId
'                                xParam(5) = ObjRequest.data.delimanName
'                                xParam(6) = ObjRequest.data.delimanPhone
'                                xParam(7) = ObjRequest.data.delimanVehicle
'                                xParam(8) = ObjRequest.data.orderTrackingUrl

'                                WService.LiveTrackingUpdate(MCon, xParam)

'                                'Insert tracking HD2 jika sudah ada status HDV IPP
'                                Query = "Select count(TrackNum) From TrackingHistory Where TrackNum=@TrackNumIPP And Status='HDV'"
'                                SqlParam = New Dictionary(Of String, String)
'                                SqlParam.Add("@TrackNumIPP", TrackNumIPP)
'                                If ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam) > 0 Then
'                                    Query = " Insert Into Tracking ("
'                                    Query &= " TrackNum, Status, TrackTime, TrackUser, TrackUserID"
'                                    Query &= " , Company1, Opr1, OprID1, OprName1"
'                                    Query &= " , Company2, Opr2, OprID2, OprName2"
'                                    Query &= " , AddInfo, TrxPartner, ForCustomer, ForCustomerEng, ForCustomerIna"
'                                    Query &= " )"
'                                    Query &= " select TrackNum, 'HD2', @TrackTime, TrackUser, TrackUserID"
'                                    Query &= " , Company1, Opr1, OprID1, OprName1"
'                                    Query &= " , '', Opr2, @delimanId, @delimanName"
'                                    Query &= " , AddInfo, TrxPartner, ForCustomer, ForCustomerEng, ForCustomerIna"
'                                    Query &= " From Tracking Where TrackNum=@TrackNumIPP And Status='HDV'"
'                                    SqlParam = New Dictionary(Of String, String)
'                                    SqlParam.Add("@TrackNumIPP", TrackNumIPP)
'                                    SqlParam.Add("@delimanId", ObjRequest.data.delimanId)
'                                    SqlParam.Add("@delimanName", ObjRequest.data.delimanName)
'                                    SqlParam.Add("@TrackTime", TrackTime)
'                                    ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

'                                    If CustomAccountLazada <> "" And CustomAccountLazada.Contains(AccountIPP) Then

'                                        'Dim WServiceLazada As New ServiceLazada

'                                        Dim wsParam(12) As Object
'                                        wsParam(0) = AccountIPP
'                                        wsParam(1) = TrackNumIPP
'                                        wsParam(2) = ObjRequest.data.salesOrderId.Trim.ToUpper
'                                        wsParam(3) = TrackTime
'                                        wsParam(4) = ""
'                                        wsParam(5) = ""
'                                        wsParam(6) = ""
'                                        wsParam(7) = ""
'                                        wsParam(8) = ""
'                                        wsParam(9) = ""
'                                        wsParam(10) = ""
'                                        wsParam(11) = ""
'                                        wsParam(12) = User

'                                        'WServiceLazada.InsertPushTrackingWI2(AppName, AppVersion, wsParam)

'                                        ReDim wsParam(12)
'                                        wsParam(0) = AccountIPP
'                                        wsParam(1) = TrackNumIPP
'                                        wsParam(2) = ObjRequest.data.salesOrderId.Trim.ToUpper
'                                        wsParam(3) = Format(CDate(TrackTime).AddSeconds(3), "yyyy-MM-dd HH:mm:ss")
'                                        wsParam(4) = ""
'                                        wsParam(5) = ""
'                                        wsParam(6) = ""
'                                        wsParam(7) = ""
'                                        wsParam(8) = "DLM"
'                                        wsParam(9) = "DLM"
'                                        wsParam(10) = ObjRequest.data.delimanId
'                                        wsParam(11) = ObjRequest.data.delimanName
'                                        wsParam(12) = User

'                                        'WServiceLazada.InsertPushTrackingHD2(AppName, AppVersion, wsParam)

'                                    End If

'                                    If CustomAccountLazadaReverse <> "" And CustomAccountLazadaReverse.Contains(AccountIPP) Then

'                                        'Dim WServiceLazada As New ServiceLazada

'                                        Dim wsParam(12) As Object
'                                        wsParam(0) = AccountIPP
'                                        wsParam(1) = TrackNumIPP
'                                        wsParam(2) = ObjRequest.data.salesOrderId.Trim.ToUpper
'                                        wsParam(3) = TrackTime
'                                        wsParam(4) = ""
'                                        wsParam(5) = ""
'                                        wsParam(6) = ""
'                                        wsParam(7) = ""
'                                        wsParam(8) = ""
'                                        wsParam(9) = ""
'                                        wsParam(10) = ""
'                                        wsParam(11) = ""
'                                        wsParam(12) = User

'                                        'WServiceLazada.InsertPushTrackingWI2(AppName, AppVersion, wsParam)

'                                        ReDim wsParam(12)
'                                        wsParam(0) = AccountIPP
'                                        wsParam(1) = TrackNumIPP
'                                        wsParam(2) = ObjRequest.data.salesOrderId.Trim.ToUpper
'                                        wsParam(3) = Format(CDate(TrackTime).AddSeconds(3), "yyyy-MM-dd HH:mm:ss")
'                                        wsParam(4) = ""
'                                        wsParam(5) = ""
'                                        wsParam(6) = ""
'                                        wsParam(7) = ""
'                                        wsParam(8) = "DLM"
'                                        wsParam(9) = "DLM"
'                                        wsParam(10) = ObjRequest.data.delimanId
'                                        wsParam(11) = ObjRequest.data.delimanName
'                                        wsParam(12) = User

'                                        'WServiceLazada.InsertPushTrackingHD2(AppName, AppVersion, wsParam)

'                                    End If

'                                End If

'                            End If
'                        Catch ex As Exception
'                        End Try

'                    End If


'                    If AlreadyOOCFLR = False And StatusIPP = "HDS" Then
'                        Query = "Update AutoOrderThirdPartyTransaction"
'                        Query &= " Set DriverInfo = '" & ("" & ObjRequest.data.deliveryPow).ToString & "'" 'foto bukti serah terima
'                        Query &= " , DriverName = '" & ObjRequest.data.delimanName & "'"
'                        Query &= " , DriverPhone = '" & ObjRequest.data.delimanPhone & "'"
'                        Query &= " , DriverId = '" & ObjRequest.data.delimanId & "'"
'                        Query &= " , DriverVehicleNo = '" & ObjRequest.data.delimanVehicle & "'"
'                        Query &= " , DriverCompany = '" & DriverCompany & "'"
'                        Query &= " , ReceiverName = '" & ObjRequest.data.receiverName & "'"
'                        Query &= " , UpdTime2 = now(), UpdUser2 = 'UPD-HDS'"
'                        Query &= " Where ThirdParty = '" & TrackUserId & "'"
'                        Query &= " And dOrderNo = '" & ObjRequest.data.orderId & "'"
'                        Query &= " And ReceiverName = ''"
'                        ObjSQL.ExecNonQuery(MCon, Query)
'                    End If


'                    If CancelType.ToUpper = "KEEP" Then
'                        Query = "Insert Into AutoOrderThirdPartyKeep ("
'                        Query &= " Account, ThirdParty, TrackNum, OrderNo, dTrackNum, dOrderNo, ServiceType, AddInfo"
'                        Query &= " , SourceType"
'                        Query &= " , Status, AddTime, AddUser"
'                        Query &= " ) Select"
'                        Query &= " Account, ThirdParty, TrackNum, OrderNo, dTrackNum, dOrderNo, ServiceType, AddInfo"
'                        Query &= " , '" & "KEP" & "'"
'                        Query &= " , '0', now(), 'NEW'"
'                        Query &= " From AutoOrderThirdPartyTransaction"
'                        Query &= " Where ThirdParty = '" & TrackUserId & "'"
'                        Query &= " And dOrderNo = '" & ObjRequest.data.orderId & "'"
'                        ObjSQL.ExecNonQuery(MCon, Query)
'                    End If 'dari If CancelType.ToUpper = "KEEP"


'                    If ObjRequest.data.status = "FAILED" Then 'fallback
'                        Query = "Insert Into AutoOrderThirdPartyKeep ("
'                        Query &= " Account, ThirdParty, TrackNum, OrderNo, dTrackNum, dOrderNo, ServiceType, AddInfo"
'                        Query &= " , SourceType"
'                        Query &= " , Status, AddTime, AddUser"
'                        Query &= " ) Select"
'                        Query &= " Account, ThirdParty, TrackNum, OrderNo, dTrackNum, dOrderNo, ServiceType, AddInfo"
'                        Query &= " , '" & "FLB" & "'"
'                        Query &= " , '0', now(), 'NEW'"
'                        Query &= " From AutoOrderThirdPartyTransaction"
'                        Query &= " Where ThirdParty = '" & TrackUserId & "'"
'                        Query &= " And dOrderNo = '" & ObjRequest.data.orderId & "'"
'                        ObjSQL.ExecNonQuery(MCon, Query)
'                    End If 'dari If ObjRequest.data.status = "FAILED"


'                    If CancelType.ToUpper = "CANCEL_BO" Then
'                        'pembatalan dari sistem BackOffice (BO) DMS
'                        'status IPP otomatis ikut FLR
'                        'kecuali untuk kiriman PAKET

'                        Query = "Select count(d.TrackNum)"
'                        Query &= " From AutoOrderThirdPartyTransaction a"
'                        Query &= " Inner Join TransactionDeliveryInfo d on (d.TrackNum = a.TrackNum and d.OrderType = 'PAKET')"
'                        Query &= " Where a.ThirdParty = '" & TrackUserId & "'"
'                        Query &= " And a.dOrderNo = '" & ObjRequest.data.orderId & "'"
'                        If ObjSQL.ExecScalar(MCon, Query) < 1 Then

'                            Dim CancelReason As String = ""
'                            Try
'                                CancelReason = ("" & ObjRequest.data.cancelCode).ToString.Trim.ToUpper
'                                'CancelReason &= "|" & ("" & ObjRequest.data.cancelReason).ToString.Trim.ToUpper
'                                CancelReason &= " - " & ("" & ObjRequest.data.cancelReason).ToString.Trim.ToUpper
'                            Catch ex As Exception
'                            End Try

'                            Query = "Insert Into Tracking ("
'                            Query &= " TrackNum, `Status`"
'                            Query &= " , TrackTime, TrackUser, TrackUserID"
'                            Query &= " , AddInfo"
'                            Query &= " ) Select"
'                            Query &= " a.TrackNum, 'FLR'"
'                            Query &= " , @TrackTime, 'SYS', 'AFL'"
'                            Query &= " , @CancelReason"
'                            Query &= " From AutoOrderThirdPartyTransaction a"
'                            Query &= " Inner Join `Transaction` t on (t.TrackNum = a.TrackNum)"
'                            Query &= " Where a.ThirdParty = @TrackUserId"
'                            Query &= " And a.dOrderNo = @orderId"

'                            SqlParam = New Dictionary(Of String, String)
'                            SqlParam.Add("@TrackTime", TrackTime)
'                            SqlParam.Add("@CancelReason", CancelReason)
'                            SqlParam.Add("@TrackUserId", TrackUserId)
'                            SqlParam.Add("@orderId", ObjRequest.data.orderId)

'                            ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

'                        End If

'                    End If 'dari If CancelType.ToUpper = "CANCEL_BO"

'SkipNext:

'                    If ErrorMessage = "" Then
'                        Query = "Update AutoOrderCallbackTracking"
'                        Query &= " Set `Status` = '1'"
'                        Query &= " , UpdTime = now(), UpdUser = 'INSERTAUTOORDER'"
'                        Query &= " Where TrackId = '" & TrackId & "'"
'                    Else
'                        Query = "Update AutoOrderCallbackTracking"
'                        Query &= " Set Retry = Retry + 1, LastError = '" & ErrorMessage & "'"
'                        Query &= " , UpdTime = now()"
'                        Query &= " Where TrackId = '" & TrackId & "'"
'                    End If
'                    ObjSQL.ExecNonQuery(MCon, Query)

'                    e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", TrackId & " " & (ErrorMessage = "").ToString, Keyword)

'                Catch ex As Exception

'                    e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, "", TrackId)

'                    Query = "Update AutoOrderCallbackTracking"
'                    Query &= " Set Retry = Retry + 1, LastError = '" & Left(WService.StrSQLRemoveChar(ex.Message), 500) & "'"
'                    Query &= " , UpdTime = now()"
'                    Query &= " Where TrackId = '" & TrackId & "'"
'                    ObjSQL.ExecNonQuery(MCon, Query)

'                End Try

'            Next

'            Process = True

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "SELESAI " & Process.ToString, Keyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'    End Sub

'    Private Function CallbackTrackingMappingStatus(ByVal StatusOri As String, ByRef PrevStatus As String) As String
'        Dim Result As String = ""

'        PrevStatus = ""

'        Select Case StatusOri.ToUpper
'            Case "PICKING_UP"
'                Result = "PUQ"

'            Case "IN_DELIVERY"
'                Result = "HDV"
'                PrevStatus = "PICKING_UP"

'            Case "COMPLETED"
'                Result = "HDS"
'                PrevStatus = "IN_DELIVERY"

'            Case "RETURN_COMPLETED"
'                Result = "CRI"
'                PrevStatus = "COMPLETED"

'                'Case "PAYMENT_COMPLETED"
'                '    Result = "CPY"
'                '    PrevStatus = "COMPLETED"
'                '#DEV3PLDELIVERYMAN

'            Case "CANCELED"
'                Result = "OOC"

'            Case "FAILED"
'                Result = "FLR"

'            Case "KEEP"
'                Result = "FLR"

'            Case "CANCEL_BO"
'                'pembatalan dari sistem BackOffice (BO) DMS
'                Result = "FLR"

'            Case "NEAR_DESTINATION"
'                Result = "NRD"

'            Case "READY_TO_DELIVER"
'                Result = "PUL" 'at pickup location

'        End Select

'        Return Result

'    End Function

'    Private Function CallbackTrackingCancelType(ByVal StatusOri As String, ByVal CancelCode As String) As String

'        Dim Result As String = ""

'        If StatusOri.ToUpper = "CANCELED" Then
'            'C001|Alamat tidak ditemukan'
'            'C002|Konsumen tidak pesan'
'            'C004|Konsumen tidak di tempat'

'            'C003|Deliveryman sakit/berhalangan'
'            'C007|Kendaraan Bermasalah

'            'C005|Pesanan sudah dibatalkan oleh toko
'            'C006|Pesanan double masuk DMS
'            'C010|Pesanan batal (mobile)

'            'C011|Pesanan sudah dibatalkan oleh toko (system)

'            'If CancelCode.Trim.ToUpper = "C003" Then
'            '    Result = "KEEP"
'            'End If

'            Select Case CancelCode.Trim.ToUpper
'                Case "C003", "C007"
'                    Result = "KEEP"

'                Case "C005", "C006", "C010", "C011"
'                    'pembatalan dari sistem BackOffice (BO) DMS
'                    Result = "CANCEL_BO"

'                Case Else

'            End Select
'        End If

'        Result = Result.ToUpper

'        Return Result

'    End Function


'    <WebMethod()>
'    Public Function CallbackTrackingLastMile(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "CallbackTrackingLastMile"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString

'            Dim Request As String = Param(2).ToString
'            Request = Request.Trim

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If Request = "" Then
'                Result(1) = "Input request"
'                GoTo Skip
'            End If

'            Dim Keyword As String = Format(Date.Now, "yyMMddHHmmss")

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "MULAI " & Request, Keyword)


'            Dim ObjRequest As cCallbackTrackingReq = JsonConvert.DeserializeObject(Of cCallbackTrackingReq)(Request)

'            If ObjRequest Is Nothing Then
'                Result(1) = "Gagal convert parameter request"
'                GoTo Skip
'            End If


'            Dim ObjSQL As New ClsSQL


'            Dim StatusTime As String = ""
'            Try
'                StatusTime = Format(CDate(ObjRequest.data.trackDate), "yyyy-MM-dd HH:mm:ss")
'            Catch ex As Exception
'                StatusTime = ObjSQL.ExecScalar(MCon, "select cast(now() as char)").ToString
'            End Try

'            Query = "Select TrackUserId"
'            Query &= " From ExpAutoOrderTrackingHistory"
'            Query &= " Where TrackNum = '" & ObjRequest.data.orderId & "'"
'            Query &= " And TrackUserId in (" & CustomAccountExpeditionDMSMotor & "," & CustomAccountExpeditionDMSMobil & ")"
'            Dim TrackUserId As String = ObjSQL.ExecScalar(MCon, Query).ToString


'            Request = JsonConvert.SerializeObject(ObjRequest)


'            'insert ke table antrian
'            Query = "Insert into ExpAutoOrderCallbackTracking ("
'            Query &= " ThirdParty, Parameter"
'            Query &= " , TrackNum, TrackStatus"
'            Query &= " , AddTime, AddUser"
'            Query &= " ) values ("
'            Query &= " '" & TrackUserId & "', '" & Request.Replace("'", "") & "'"
'            Query &= " , '" & ObjRequest.data.orderId & "', '" & ObjRequest.data.status & "'"
'            Query &= " , '" & StatusTime & "', 'NEW'"
'            Query &= " )"
'            If ObjSQL.ExecNonQuery(MCon, Query) = False Then
'                Result(1) = "Gagal insert antrian " & ObjRequest.data.orderId
'                GoTo Skip
'            End If

'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = ""

'            Process = True

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "SELESAI " & Process.ToString, Keyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function

'    <WebMethod()>
'    Public Sub CallbackTrackingProcessLastMile(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object())

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "CallbackTrackingProcessLastMile"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            AppName = wsAppName
'            AppVersion = wsAppVersion

'            Dim Keyword As String = Format(Date.Now, "yyMMddHHmmss")

'            e.DebugLog(MCon, AppName, AppVersion, Method, wsAppName, "MULAI", Keyword)

'            Dim ObjSQL As New ClsSQL

'            Query = "Select ThirdParty, TrackId, TrackNum, TrackStatus"
'            Query &= " , cast(Parameter as char) as Request"
'            Query &= " , cast(AddTime as char) as TrackTime"
'            Query &= " From ExpAutoOrderCallbackTracking"
'            Query &= " Where `Status` = '0'"
'            'Query &= " And TrackStatus in ('COMPLETED', 'CANCELED', 'CANCEL_BO')"
'            Query &= " And ThirdParty in (" & CustomAccountExpeditionDMSMotor & "," & CustomAccountExpeditionDMSMobil & ")"
'            Query &= " Order By AddTime, Retry"
'            Query &= " Limit 100"
'            Dim dtProcess As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)

'            For p As Integer = 0 To dtProcess.Rows.Count - 1

'                Dim ThirdParty As String = dtProcess.Rows(p).Item("ThirdParty").ToString
'                Dim TrackId As String = dtProcess.Rows(p).Item("TrackId").ToString
'                Dim TrackNum As String = dtProcess.Rows(p).Item("TrackNum").ToString
'                Dim TrackStatus As String = dtProcess.Rows(p).Item("TrackStatus").ToString

'                Try
'                    Dim Request As String = dtProcess.Rows(p).Item("Request").ToString

'                    Dim ErrorMessage As String = ""

'                    Dim ObjRequest As cCallbackTrackingReq = JsonConvert.DeserializeObject(Of cCallbackTrackingReq)(Request)

'                    If ObjRequest Is Nothing Then
'                        ErrorMessage = "Gagal convert parameter request " & TrackId
'                        GoTo SkipNext
'                    End If

'                    'validasi
'                    Dim StatusIPP As String = CallbackTrackingMappingStatusLastMile(TrackStatus)
'                    If StatusIPP = "" Then
'                        ErrorMessage = "Gagal mapping status " & TrackStatus & " " & TrackId
'                        GoTo SkipNext
'                    End If

'                    Query = "Select count(TrackNum)"
'                    Query &= " From ExpAutoOrderTrackingHistory"
'                    Query &= " Where TrackNum = @TrackNum"
'                    Query &= " And TrackUserId in (" & CustomAccountExpeditionDMSMotor & "," & CustomAccountExpeditionDMSMobil & ")"

'                    SqlParam = New Dictionary(Of String, String)
'                    SqlParam.Add("@TrackNum", TrackNum)

'                    If ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam) < 1 Then
'                        ErrorMessage = "Nomor " & TrackNum & " tidak ditemukan " & TrackId
'                        GoTo SkipNext
'                    End If

'                    Dim TrackTime As String = dtProcess.Rows(p).Item("TrackTime").ToString

'                    Query = "Select count(TrackNum)"
'                    Query &= " From ExpAutoOrderTracking"
'                    Query &= " Where TrackNum = @TrackNum"
'                    Query &= " And TrackUserId in (" & CustomAccountExpeditionDMSMotor & "," & CustomAccountExpeditionDMSMobil & ")"
'                    Query &= " And `Status` = @StatusIPP"
'                    Query &= " And TrackTime = @TrackTime"

'                    SqlParam = New Dictionary(Of String, String)
'                    SqlParam.Add("@TrackNum", TrackNum)
'                    SqlParam.Add("@StatusIPP", StatusIPP)
'                    SqlParam.Add("@TrackTime", TrackTime)

'                    If ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam) > 0 Then
'                        ErrorMessage = "" 'sudah pernah dilakukan
'                        GoTo SkipNext
'                    End If

'                    Query = "Select xaoh.TrackUserId, t.Weight"
'                    Query &= " From ExpAutoOrderTrackingHistory xaoh"
'                    Query &= " Join ExpAutoOrderTrx xaot on xaoh.TrackNum=xaot.DeliveryId"
'                    Query &= " Join Transaction t on xaot.TrackNum=t.TrackNum"
'                    Query &= " Where xaoh.TrackNum = '" & ObjRequest.data.orderId & "'"
'                    Query &= " And xaoh.TrackUserId in (" & CustomAccountExpeditionDMSMotor & "," & CustomAccountExpeditionDMSMobil & ")"
'                    Dim TrackUserId As String = ""
'                    Dim TransactionWeight As String = ""
'                    Try
'                        Dim dtTemp As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)
'                        TrackUserId = dtTemp.Rows(0).Item("TrackUserId").ToString()
'                        TransactionWeight = dtTemp.Rows(0).Item("Weight").ToString()
'                    Catch ex As Exception
'                        TrackUserId = ""
'                        TransactionWeight = ""

'                        ErrorMessage = "Invalid " & ObjRequest.data.orderId & " " & Query
'                        GoTo SkipNext
'                    End Try


'                    ObjRequest.data.orderTrackingUrl = ("" & ObjRequest.data.orderTrackingUrl)

'                    ObjRequest.data.delimanId = ("" & ObjRequest.data.delimanId)
'                    ObjRequest.data.delimanId = Left(ObjRequest.data.delimanId, 45)

'                    ObjRequest.data.delimanName = ("" & ObjRequest.data.delimanName)
'                    ObjRequest.data.delimanName = WService.StrSQLRemoveChar(Left(ObjRequest.data.delimanName, 100))

'                    ObjRequest.data.delimanVehicle = ("" & ObjRequest.data.delimanVehicle)
'                    ObjRequest.data.delimanVehicle = Left(ObjRequest.data.delimanVehicle, 45)

'                    Dim DriverCompany As String = ""
'                    Try
'                        DriverCompany = ObjRequest.data.delimanCompany.Trim.ToUpper
'                    Catch ex As Exception
'                        DriverCompany = ""
'                    End Try

'                    ObjRequest.data.receiverName = ("" & ObjRequest.data.receiverName)
'                    ObjRequest.data.receiverName = WService.StrSQLRemoveChar(Left(ObjRequest.data.receiverName, 100))

'                    ObjRequest.data.receiverRelationship = ("" & ObjRequest.data.receiverRelationship)
'                    ObjRequest.data.receiverRelationship = Left(ObjRequest.data.receiverRelationship, 45)


'                    'insert data

'                    Query = "Insert Into ExpAutoOrderTracking ("
'                    Query &= " TrackNum, `Status`, TrxPartner"
'                    Query &= " , TrackTime, TrackUser, TrackUserID"
'                    Query &= " , ForCustomer"
'                    If (StatusIPP = "RTS") And ObjRequest.data.delimanName <> "" Then
'                        Query &= " , OprID1, OprName1" '-- nama pengirim terakhir, sebelum gagal kirim / gagal proses
'                    End If
'                    If StatusIPP = "HDS" Then
'                        Query &= " , OprID1, OprName1" '-- nama pengirim
'                        Query &= " , OprID2, OprName2" '-- nama penerima
'                    End If

'                    If StatusIPP = "HDS" Or StatusIPP = "RTS" Then
'                        Query &= " , AddInfo" '-- foto/bukti/alasan proses
'                    End If

'                    Query &= " ) values ("
'                    Query &= " '" & ObjRequest.data.orderId & "', '" & StatusIPP & "', '" & ObjRequest.data.status.ToUpper & "'"
'                    Query &= " , '" & TrackTime & "', 'DELIMAN', '" & TrackUserId & "'"

'                    If StatusIPP = "HDS" Then
'                        Query &= " , '1'" 'ForCustomer
'                    Else
'                        Query &= " , '0'"
'                    End If
'                    If (StatusIPP = "RTS") And ObjRequest.data.delimanName <> "" Then
'                        Query &= " , '" & ObjRequest.data.delimanId & "', '" & ObjRequest.data.delimanName & "'" '-- nama pengirim terakhir, sebelum gagal kirim / gagal proses
'                    End If
'                    If StatusIPP = "HDS" Then
'                        Query &= " , '" & ObjRequest.data.delimanId & "', '" & ObjRequest.data.delimanName & "'" '-- nama pengirim
'                        Query &= " , '" & ObjRequest.data.receiverRelationship & "', '" & ObjRequest.data.receiverName & "'" '-- nama penerima
'                    End If


'                    Dim Reason As String = ""
'                    Try
'                        Reason = ("" & ObjRequest.data.failReason).ToString
'                    Catch ex As Exception
'                    End Try
'                    Try
'                        Reason = ("" & ObjRequest.data.cancelCode).ToString & "|" & ("" & ObjRequest.data.cancelReason).ToString
'                    Catch ex As Exception
'                    End Try

'                    If StatusIPP = "HDS" Then
'                        Query &= " , '" & ("" & ObjRequest.data.deliveryPow).ToString & "'" '-- bukti diterima
'                    ElseIf StatusIPP = "RTS" Then
'                        Query &= " , '" & Reason & "'" '-- alasan gagal proses
'                    End If

'                    Query &= " )"
'                    If ObjSQL.ExecNonQuery(MCon, Query) = False Then
'                        ErrorMessage = "Gagal insert ExpAutoOrderTracking " & TrackId
'                        GoTo SkipNext
'                    End If


'                    Dim DriverName As String = "DELIMAN"
'                    If StatusIPP = "HDS" Then

'                        Dim ReceiverName As String = ""
'                        Try
'                            ReceiverName = ("" & ObjRequest.data.receiverName)
'                            ReceiverName = WService.StrSQLRemoveChar(ReceiverName)
'                        Catch ex As Exception
'                            ReceiverName = ""
'                        End Try
'                        If ReceiverName.Trim = "" Then
'                            ReceiverName = "Penerima"
'                        End If
'                        ReceiverName = Left(WService.StrSQLRemoveChar(ReceiverName), 100)

'                        Dim HdsPhoto As String = ""
'                        Try
'                            HdsPhoto = Left(ObjRequest.data.deliveryPow, 500)
'                        Catch ex As Exception
'                            HdsPhoto = ""
'                        End Try


'                        SqlParam = New Dictionary(Of String, String)

'                        Query = "Update ExpAutoOrderTrx"
'                        Query &= " Set DriverName = @DriverName"
'                        'Query &= " , DriverPhone = '" & Strings.Left(DriverPhone, 100) & "'"
'                        Query &= " , ReceiverName = @ReceiverName"
'                        If HdsPhoto <> "" Then
'                            Query &= " , DriverInfo = @HdsPhoto"
'                            SqlParam.Add("@HdsPhoto", HdsPhoto)
'                        End If
'                        If TransactionWeight <> "" Then
'                            Query &= " , Weight = @Weight"
'                            SqlParam.Add("@Weight", TransactionWeight)
'                        End If
'                        Query &= " , UpdTime2 = now(), UpdUser2 = 'UPD-HDS'"
'                        Query &= " Where ThirdParty in (" & CustomAccountExpeditionDMSMotor & "," & CustomAccountExpeditionDMSMobil & ")"
'                        Query &= " And DeliveryId = @TrackNum"

'                        SqlParam.Add("@DriverName", Strings.Left(DriverName, 100))
'                        SqlParam.Add("@ReceiverName", Strings.Left(ReceiverName, 100))
'                        SqlParam.Add("@TrackNum", TrackNum)

'                        ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)
'                    End If


'                    If StatusIPP = "RTS" Then

'                        Dim DeliveryCost As Double = 1

'                        Dim RtsParam(8) As Object
'                        RtsParam(0) = User
'                        RtsParam(1) = Password
'                        RtsParam(2) = "" 'TrackNum
'                        RtsParam(3) = TrackUserId 'oExpedition
'                        RtsParam(4) = TrackNum 'oAWB
'                        RtsParam(5) = TransactionWeight 'oWeight
'                        RtsParam(6) = Format(CDate(IIf(TrackTime <> "", TrackTime, Date.Now)), "yyyy-MM-dd HH:mm:ss") 'oAwbDate
'                        RtsParam(7) = DeliveryCost 'oAwbDcost
'                        RtsParam(8) = wsAppName & "_RTS"
'                        Try
'                            WService.SetOtherExpeditionRTS(wsAppName, wsAppVersion, RtsParam)
'                        Catch ex As Exception
'                            e.ErrorLog(MCon, AppName, AppVersion, Method, wsAppName, ex, "Gagal SetOtherExpeditionRTS", TrackId)
'                        End Try

'                    End If 'dari If StatusIPP = "RTS"

'SkipNext:

'                    SqlParam = New Dictionary(Of String, String)

'                    If ErrorMessage = "" Then
'                        Query = "Update ExpAutoOrderCallbackTracking"
'                        Query &= " Set `Status` = '1'"
'                        Query &= " , UpdTime = now(), UpdUser = 'INSERTExpAutoOrder'"
'                        Query &= " Where TrackId = @TrackId"
'                    Else
'                        Query = "Update ExpAutoOrderCallbackTracking"
'                        Query &= " Set Retry = Retry + 1, LastError = @ErrorMessage"
'                        Query &= " , UpdTime = now()"
'                        Query &= " Where TrackId = @TrackId"

'                        SqlParam.Add("@ErrorMessage", ErrorMessage)

'                    End If

'                    SqlParam.Add("@TrackId", TrackId)
'                    ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

'                    e.DebugLog(MCon, AppName, AppVersion, Method, wsAppName, TrackId & " " & (ErrorMessage = "").ToString, Keyword)

'                Catch ex As Exception

'                    e.ErrorLog(MCon, AppName, AppVersion, Method, wsAppName, ex, "", TrackId)

'                    Query = "Update ExpAutoOrderCallbackTracking"
'                    Query &= " Set Retry = Retry + 1, LastError = @LastError"
'                    Query &= " , UpdTime = now()"
'                    Query &= " Where TrackId = @TrackId"

'                    SqlParam = New Dictionary(Of String, String)
'                    SqlParam.Add("@LastError", Left(WService.StrSQLRemoveChar(ex.Message), 500))
'                    SqlParam.Add("@TrackId", TrackId)
'                    ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

'                End Try

'            Next

'            Process = True

'Skip:

'            e.DebugLog(MCon, AppName, AppVersion, Method, wsAppName, "SELESAI " & Process.ToString, Keyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, AppName, AppVersion, Method, wsAppName, ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'    End Sub

'    Private Function CallbackTrackingMappingStatusLastMile(ByVal StatusOri As String) As String
'        Dim Result As String = ""

'        Select Case StatusOri.ToUpper
'            Case "PICKING_UP"
'                Result = "PUQ"

'            Case "IN_DELIVERY"
'                Result = "HDV"

'            Case "COMPLETED"
'                Result = "HDS"

'            Case "CANCELED"
'                Result = "RTS"

'            Case "FAILED"
'                Result = "FLR"

'            Case "KEEP"
'                Result = "FLR"

'            Case "CANCEL_BO"
'                'pembatalan dari sistem BackOffice (BO) DMS
'                Result = "RTS"

'            Case "READY_TO_DELIVER"
'                Result = "PUL" 'at pickup location

'        End Select

'        Return Result

'    End Function


'    Private Function FormatName(ByVal StrOri As String) As String

'        On Error Resume Next

'        Dim Result As String = ""

'        For i As Integer = 0 To StrOri.Length - 1
'            If IsNumeric(StrOri(i)) Or ("ABCDEFGHIJKLMNOPQRSTUVWXYZ").Contains(StrOri(i).ToString.ToUpper) Or StrOri(i) = " " Then
'                Result &= StrOri(i)
'            Else
'                Result &= " "
'            End If
'        Next

'        Result = Result.Replace("  ", " ").Replace("  ", " ")

'        Return Result

'    End Function

'    Private Function FormatPhone(ByVal StrOri As String) As String

'        On Error Resume Next

'        Dim Result As String = ""

'        For i As Integer = 0 To StrOri.Length - 1
'            If IsNumeric(StrOri(i)) Then
'                Result &= StrOri(i)
'            End If
'        Next

'        If Result.StartsWith("0") Then
'            Result = Strings.Right(Result, Result.Length - 1)
'            Result = "62" & Result
'        End If

'        Return Result

'    End Function


'    <WebMethod()>
'    Public Function GetShStore(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object

'        Dim Method As String = "GetShStore"

'        Dim Result As Object() = CreateResult()
'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim StoreCode As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString

'            StoreCode = Param(2).ToString

'            Dim OrderSource As String = ""
'            Try
'                OrderSource = Param(3).ToString.Trim.ToUpper
'            Catch ex As Exception
'                OrderSource = ""
'            End Try
'            If OrderSource = "" Then
'                OrderSource = "INDOMARET"
'            End If

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If OrderSource = "INDOGROSIR" Then
'                Query = "Select concat(i.PrefixDmsCode, i.Code) as Code, i.Name, i.Address, i.PostalCode, i.Phone, i.Latitude, i.Longitude"
'                Query &= " , i.NumCode as DcCode, i.Code as DcName"
'                Query &= " From IndogrosirStore i"
'                Query &= " Where i.NumCode = '" & StoreCode & "'"
'            Else
'                'INDOMARET
'                Query = "Select i.Code, i.Name, i.Address, i.PostalCode, i.Phone, i.Latitude, i.Longitude"
'                Query &= " , cast(ifnull(dc.Code,'') as char) as DcCode, cast(ifnull(dc.Alias,'') as char) as DcName"
'                Query &= " From IndomaretStore_For_Report i"
'                Query &= " Left Join IndomaretDC dc on (i.DcCode = dc.Code)"
'                Query &= " Where i.Code = '" & StoreCode & "'"
'            End If

'            Dim ObjSQL As New ClsSQL

'            Dim dtStore As DataTable = ObjSQL.ExecDatatable(MCon, Query)

'            If dtStore Is Nothing Then
'                Result(1) = "Gagal get Data Toko"
'                GoTo Skip
'            End If

'            If dtStore.Rows.Count < 1 Then
'                Result(1) = "Toko " & StoreCode & " tidak ditemukan"
'                GoTo Skip
'            End If


'            'pakai nomor telp yang didaftarkan oleh IPP
'            Query = "Select trim(cast(case when ifnull(PhonePIC1,'') <> '' then PhonePIC1 else ifnull(PhonePIC2,'') end as char)) as PhoneExt"
'            Query &= " From IndomaretStore_ExtInfo"
'            Query &= " Where Code = '" & StoreCode & "'"
'            Dim PhoneExt As String = ("" & ObjSQL.ExecScalar(MCon, Query))
'            If PhoneExt <> "" Then
'                dtStore.Rows(0).Item("Phone") = PhoneExt
'                dtStore.AcceptChanges()
'            End If


'            ReDim Result(10)
'            Result(0) = "0"
'            Result(1) = ""
'            Result(2) = dtStore.Rows(0).Item("Code").ToString.ToUpper
'            Result(3) = dtStore.Rows(0).Item("Name").ToString.ToUpper
'            Result(4) = dtStore.Rows(0).Item("Address").ToString.ToUpper
'            Result(5) = dtStore.Rows(0).Item("PostalCode").ToString.ToUpper
'            Result(6) = dtStore.Rows(0).Item("Phone").ToString.ToUpper
'            Result(7) = dtStore.Rows(0).Item("Latitude").ToString.ToUpper
'            Result(8) = dtStore.Rows(0).Item("Longitude").ToString.ToUpper
'            Result(9) = dtStore.Rows(0).Item("DcCode").ToString.ToUpper
'            Result(10) = dtStore.Rows(0).Item("DcName").ToString.ToUpper
'            '2 = code, 3 = name, 4 = address, 5 = postalcode, 6 = phone, 7 = latitude, 8 = longitude, 9 = DcCode, 10 = DcName

'Skip:

'            'e.DebugLog( AppName, AppVersion, Method, "WServiceDeliMan", "FINISH", StoreCode)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "", ex, "", StoreCode)

'            Result(1) = ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function

'    <WebMethod()>
'    Public Function GetLatLonByAddressPostalCode(ByVal Param As Object()) As Object()

'        Dim Result As Object() = CreateResult()

'        Dim Method As String = "GetLatLonByAddressPostalCode"

'        Dim MCon As MySqlConnection

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString
'            Dim Address As String = Param(2).ToString.Trim
'            Dim PostalCode As String = Param(3).ToString.Trim

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            Dim dt As New DataTable
'            dt.Columns.Add("REQADDRESS")
'            dt.Columns.Add("REQPOSTALCODE")
'            dt.Columns.Add("REQLATITUDE")
'            dt.Columns.Add("REQLONGITUDE")

'            Dim dr As DataRow
'            dr = dt.NewRow

'            dr.Item("REQADDRESS") = Address
'            dr.Item("REQPOSTALCODE") = PostalCode
'            dr.Item("REQLATITUDE") = ""
'            dr.Item("REQLONGITUDE") = ""

'            dt.Rows.Add(dr)

'            Dim ObjFungsi As New ClsFungsi
'            Dim ErrorMessage As String = ""
'            Dim dtResult As DataTable = ObjFungsi.GoogleGeocodingLocation(dt, ErrorMessage)

'            If dtResult Is Nothing Then
'                Result(1) = "dtResult is nothing"
'                If ErrorMessage <> "" Then
'                    Result(1) &= " " & ErrorMessage
'                End If
'                GoTo Skip
'            End If

'            If dtResult.Rows.Count < 1 Then
'                Result(1) = "dtResult no row"
'                If ErrorMessage <> "" Then
'                    Result(1) &= " " & ErrorMessage
'                End If
'                GoTo Skip
'            End If


'            If dtResult.Rows(0).Item("Status") <> "OK" Then
'                Result(1) = "Status Pencarian : " & dtResult.Rows(0).Item("Status")
'                If ErrorMessage <> "" Then
'                    Result(1) &= " " & ErrorMessage
'                End If
'                GoTo Skip
'            End If

'            Dim Latitude As String = dtResult.Rows(0).Item("REQLATITUDE")
'            Dim Longitude As String = dtResult.Rows(0).Item("REQLONGITUDE")

'            Result(0) = "0"
'            Result(1) = ""
'            Result(2) = Latitude & "|" & Longitude

'            Process = True

'Skip:

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog("", "", Method, User, ex, "")

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try
'        End Try

'        Return Result

'    End Function


'    <WebMethod()>
'    Public Function ProcessTrackNumSTIBeforeBKO(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object() = CreateResult()

'        Dim Method As String = "ProcessTrackNumSTIBeforeBKO"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim LogKeyword As String = ""

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString
'            Dim TrackNumList As String = Param(2).ToString.Trim.ToUpper

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If TrackNumList = "" Then
'                Result(1) = "Input TrackNum"
'                GoTo Skip
'            End If

'            LogKeyword = User & " " & Format(Date.Now, "yyMMddHHmmss")
'            LogKeyword = WService.StrSQLRemoveChar(LogKeyword)

'            e.RequestLog(MCon, AppName, AppVersion, Method, User, "MULAI", LogKeyword)

'            Dim TrackNumAll As String() = TrackNumList.Split("|")

'            Dim dtTrackNumAll As New DataTable
'            dtTrackNumAll.Columns.Add("TrackNum")
'            dtTrackNumAll.Columns.Add("Status")

'            For i As Integer = 0 To TrackNumAll.Length - 1
'                If TrackNumAll(i) <> "" Then
'                    dtTrackNumAll.Rows.Add()
'                    dtTrackNumAll.Rows(dtTrackNumAll.Rows.Count - 1).Item("TrackNum") = TrackNumAll(i)
'                    dtTrackNumAll.Rows(dtTrackNumAll.Rows.Count - 1).Item("Status") = ""
'                    dtTrackNumAll.AcceptChanges()
'                End If
'            Next


'            If dtTrackNumAll.Rows.Count < 1 Then
'                Result(1) = "Input TrackNum "
'                GoTo Skip
'            End If

'            Dim MaxTrackNum As Integer = 50
'            If dtTrackNumAll.Rows.Count > MaxTrackNum Then
'                Result(1) = "Max. TrackNum " & MaxTrackNum
'                GoTo Skip
'            End If


'            Dim ObjSQL As New ClsSQL

'            Query = "Select Key1, `Value` From `Const` Where Key1 in ('CRODLMLMT','CRODLMMIN','CRODLMMAX','CRODLMTPR','CRODLMTBF')"
'            Dim dtCreateOrderDeliveryMan As DataTable = ObjSQL.ExecDatatable(MCon, Query)

'            'Batas akhir create order ke sistem DeliveryMan pada hari H, pukul ...
'            Dim CreateOrderDeliveryManHourLimitDaily As String = ""
'            Try
'                dtCreateOrderDeliveryMan.DefaultView.RowFilter = "Key1 = 'CRODLMLMT'"
'                CreateOrderDeliveryManHourLimitDaily = dtCreateOrderDeliveryMan.DefaultView(0).Item("Value").ToString
'            Catch ex As Exception
'                CreateOrderDeliveryManHourLimitDaily = ""
'            End Try
'            If CreateOrderDeliveryManHourLimitDaily = "" Then
'                CreateOrderDeliveryManHourLimitDaily = "15"
'            End If

'            'Batas bawah delivery time pada sistem DeliveryMan setiap hari, pukul ...
'            Dim CreateOrderDeliveryManMinTime As String = ""
'            Try
'                dtCreateOrderDeliveryMan.DefaultView.RowFilter = "Key1 = 'CRODLMMIN'"
'                CreateOrderDeliveryManMinTime = dtCreateOrderDeliveryMan.DefaultView(0).Item("Value").ToString
'            Catch ex As Exception
'                CreateOrderDeliveryManMinTime = ""
'            End Try
'            If CreateOrderDeliveryManMinTime = "" Then
'                CreateOrderDeliveryManMinTime = "08"
'            End If

'            'Batas atas delivery time pada sistem DeliveryMan setiap hari, pukul ...
'            Dim CreateOrderDeliveryManMaxTime As String = ""
'            Try
'                dtCreateOrderDeliveryMan.DefaultView.RowFilter = "Key1 = 'CRODLMMAX'"
'                CreateOrderDeliveryManMaxTime = dtCreateOrderDeliveryMan.DefaultView(0).Item("Value").ToString
'            Catch ex As Exception
'                CreateOrderDeliveryManMaxTime = ""
'            End Try
'            If CreateOrderDeliveryManMaxTime = "" Then
'                CreateOrderDeliveryManMaxTime = "20"
'            End If

'            'Jeda untuk antisipasi waktu proses yang diperlukan oleh sistem atau API
'            Dim CreateOrderDeliveryManProcessTime As String = ""
'            Try
'                dtCreateOrderDeliveryMan.DefaultView.RowFilter = "Key1 = 'CRODLMTPR'"
'                CreateOrderDeliveryManProcessTime = dtCreateOrderDeliveryMan.DefaultView(0).Item("Value").ToString
'            Catch ex As Exception
'                CreateOrderDeliveryManProcessTime = ""
'            End Try
'            If CreateOrderDeliveryManMaxTime = "" Then
'                CreateOrderDeliveryManProcessTime = "15"
'            End If

'            'Buffer tambahan untuk waktu proses yang diperlukan oleh sistem atau API
'            Dim CreateOrderDeliveryManProcessBuffer As String = ""
'            Try
'                dtCreateOrderDeliveryMan.DefaultView.RowFilter = "Key1 = 'CRODLMTBF'"
'                CreateOrderDeliveryManProcessBuffer = dtCreateOrderDeliveryMan.DefaultView(0).Item("Value").ToString
'            Catch ex As Exception
'                CreateOrderDeliveryManProcessBuffer = ""
'            End Try
'            If CreateOrderDeliveryManProcessBuffer = "" Then
'                CreateOrderDeliveryManProcessBuffer = "5"
'            End If


'            For i As Integer = 0 To dtTrackNumAll.Rows.Count - 1

'                Dim ErrMsg As String = ""

'                Dim TrackNum As String = dtTrackNumAll.Rows(i).Item("TrackNum").ToString

'                e.RequestLog(MCon, AppName, AppVersion, Method, User, "MULAI", LogKeyword & " " & TrackNum)

'                Query = "Select t.TrackNum"
'                Query &= " , t.ShAccount as Account"
'                Query &= " , cast(case when ifnull(t.OrderNo,'') <> '' then t.OrderNo else t.TrackNum end as char) as OrderNo"
'                Query &= " , t.CoAddress, t.CoPostalCode"
'                Query &= " , t.CoLatitude, t.CoLongitude"
'                Query &= " From `Transaction` t"
'                Query &= " Inner Join MstService v on (v.Service = t.ServiceType)"
'                Query &= " Inner Join Tracking r on (r.TrackNum = t.TrackNum and r.`Status` = 'STI')"
'                Query &= " Where t.TrackNum = '" & TrackNum & "'"
'                Query &= " And t.TrackNum = t.oTrackNum"
'                Query &= " And ucase(v.Name) Like '%DOOR'"

'                Dim dtTransaction As DataTable = ObjSQL.ExecDatatable(MCon, Query)
'                If dtTransaction Is Nothing Then
'                    ErrMsg = "Gagal query"
'                    GoTo SkipNext
'                End If

'                If dtTransaction.Rows.Count < 1 Then
'                    ErrMsg = "Tidak valid (Ditemukan . AWB Semula . Layanan Door . Status MTO)"
'                    GoTo SkipNext
'                End If

'                Query = "Select count(TrackNum)"
'                Query &= " From TransactionDeliveryInfo"
'                Query &= " Where TrackNum = '" & TrackNum & "'"
'                If ObjSQL.ExecScalar(MCon, Query) > 0 Then
'                    ErrMsg = "Sudah ada DeliveryInfo"
'                    GoTo SkipNext
'                End If


'                '=== mirip dengan proses STI awb to-Door

'                Dim Account As String = dtTransaction.Rows(0).Item("Account").ToString
'                Dim OrderNo As String = dtTransaction.Rows(0).Item("OrderNo").ToString

'                Dim DeliveryInfoPINPickup As String = ""
'                Dim DeliveryInfoPINCancel As String = ""
'                Dim DeliveryInfoPINKeep As String = ""
'                Dim DeliveryInfoPINReturn As String = ""
'                Dim DeliveryInfoPINBase As String = ""
'                Try
'                    Dim DeliveryInfoPIN As String() = WService.GenerateDeliveryInfoPIN()
'                    DeliveryInfoPINPickup = DeliveryInfoPIN(0)
'                    DeliveryInfoPINCancel = DeliveryInfoPIN(1)
'                    DeliveryInfoPINKeep = DeliveryInfoPIN(2)
'                    DeliveryInfoPINReturn = DeliveryInfoPIN(3)
'                    DeliveryInfoPINReturn = "" 'PAKET, tidak ada return bulky
'                    DeliveryInfoPINBase = DeliveryInfoPIN(4)

'                    If DeliveryInfoPINBase = "XXXXX" Then
'                        ErrMsg = "Gagal Generate PIN (DeliveryInfo)"
'                        GoTo SkipNext
'                    End If

'                Catch ex As Exception
'                    DeliveryInfoPINPickup = ""
'                    DeliveryInfoPINCancel = ""
'                    DeliveryInfoPINKeep = ""
'                    DeliveryInfoPINReturn = ""
'                    DeliveryInfoPINBase = ""
'                End Try


'                Dim SendRequestTime As String = ""
'                Dim ExpectedDeliverMinTime As String = ""
'                Dim ExpectedDeliverMaxTime As String = ""
'                Dim DeliveryInfoTimeZone As Integer = 0
'                Try
'                    'pembatasan waktu order ke sistem deliveryman (SAMEDAY)
'                    'bila lewat jam tertentu, maka dialihkan ke hari berikut
'                    Query = "Select cast(concat(CanOrderSameday, '|', RequestTime, '|', DeliveryTimeMin, '|', DeliveryTimeMax) as char) as `Value`"
'                    Query &= " From ("
'                    Query &= "   Select x.CanOrderSameday"
'                    Query &= "   , cast(now() as char) as RequestTime"
'                    Query &= "   , cast((case when x.CanOrderSameday = '1' then x.SqlNowPlus                                                      else concat(date_add(curdate(),interval 1 day), ' " & CreateOrderDeliveryManMinTime & ":00:00') end) as char) as DeliveryTimeMin"
'                    Query &= "   , cast((case when x.CanOrderSameday = '1' then concat(curdate(), ' " & CreateOrderDeliveryManMaxTime & ":00:00') else concat(date_add(curdate(),interval 1 day), ' " & CreateOrderDeliveryManMaxTime & ":00:00') end) as char) as DeliveryTimeMax"
'                    Query &= "   From ("
'                    Query &= "     Select now() as SqlNow"
'                    Query &= "     , date_add(now(), interval (" & CreateOrderDeliveryManProcessTime & " + " & CreateOrderDeliveryManProcessBuffer & ") minute) as SqlNowPlus"
'                    Query &= "     , concat(curdate(), ' " & CreateOrderDeliveryManHourLimitDaily & ":00:00') as OrderTimeLimit"
'                    Query &= "     , cast((case when date_add(now(), interval (" & CreateOrderDeliveryManProcessTime & " + " & CreateOrderDeliveryManProcessBuffer & ") minute) < concat(curdate(), ' " & CreateOrderDeliveryManHourLimitDaily & ":00:00') then '1' else '0' end) as char) as CanOrderSameday"
'                    Query &= "   )x"
'                    Query &= " )y"
'                    Dim CreateOrderDeliveryManTimeValue As String() = ObjSQL.ExecScalar(MCon, Query).ToString.Split("|")

'                    Dim DiffTimeZone As Integer = WService.SetDiffTimeZoneByPostalCode(dtTransaction.Rows(0).Item("CoPostalCode").ToString, DeliveryInfoTimeZone)

'                    SendRequestTime = CreateOrderDeliveryManTimeValue(1)
'                    Try
'                        SendRequestTime = Format(DateAdd(DateInterval.Hour, DiffTimeZone * -1, CDate(SendRequestTime)), "yyyy-MM-dd HH:mm:ss")
'                    Catch ex As Exception
'                    End Try

'                    ExpectedDeliverMinTime = CreateOrderDeliveryManTimeValue(2)
'                    Try
'                        ExpectedDeliverMinTime = Format(DateAdd(DateInterval.Hour, DiffTimeZone * -1, CDate(ExpectedDeliverMinTime)), "yyyy-MM-dd HH:mm:ss")
'                    Catch ex As Exception
'                    End Try

'                    ExpectedDeliverMaxTime = CreateOrderDeliveryManTimeValue(3)
'                    Try
'                        ExpectedDeliverMaxTime = Format(DateAdd(DateInterval.Hour, DiffTimeZone * -1, CDate(ExpectedDeliverMaxTime)), "yyyy-MM-dd HH:mm:ss")
'                    Catch ex As Exception
'                    End Try

'                Catch ex As Exception
'                    SendRequestTime = ""
'                End Try
'                If SendRequestTime = "" Then
'                    ErrMsg = "Gagal Generate Order and Delivery Time (DeliveryInfo)"
'                    GoTo SkipNext
'                End If


'                Try
'                    Dim CoNeedLatLon As Boolean = False
'                    If dtTransaction.Rows(0).Item("CoLatitude").ToString = "" _
'                    Or dtTransaction.Rows(0).Item("CoLongitude").ToString = "" Then
'                        CoNeedLatLon = True
'                    End If
'                    Try
'                        If CDbl(dtTransaction.Rows(0).Item("CoLatitude")) = 0 And CDbl(dtTransaction.Rows(0).Item("CoLongitude")) = 0 Then
'                            CoNeedLatLon = True
'                        End If
'                    Catch ex As Exception
'                    End Try

'                    If CoNeedLatLon Then
'                        Dim dtLatLon As New DataTable
'                        dtLatLon.Columns.Add("REQADDRESS")
'                        dtLatLon.Columns.Add("REQPOSTALCODE")
'                        dtLatLon.Columns.Add("REQLATITUDE")
'                        dtLatLon.Columns.Add("REQLONGITUDE")

'                        Dim dr As DataRow
'                        dr = dtLatLon.NewRow

'                        dr.Item("REQADDRESS") = dtTransaction.Rows(0).Item("CoAddress").ToString
'                        dr.Item("REQPOSTALCODE") = dtTransaction.Rows(0).Item("CoPostalCode").ToString
'                        dr.Item("REQLATITUDE") = ""
'                        dr.Item("REQLONGITUDE") = ""

'                        dtLatLon.Rows.Add(dr)

'                        Dim ObjFungsi As New ClsFungsi
'                        Dim LatLonErrorMessage As String = ""
'                        Dim dtResult As DataTable = ObjFungsi.GoogleGeocodingLocation(dtLatLon, LatLonErrorMessage)

'                        Dim ReqLatitude As String = dtResult.Rows(0).Item("REQLATITUDE")
'                        Dim ReqLongitude As String = dtResult.Rows(0).Item("REQLONGITUDE")

'                        If LatLonErrorMessage <> "" Then
'                            ErrMsg &= " " & LatLonErrorMessage

'                        ElseIf ReqLatitude <> "" And ReqLongitude <> "" Then
'                            Query = "Update `Transaction`"
'                            Query &= " Set CoLatitude = '" & ReqLatitude & "', CoLongitude = '" & ReqLongitude & "'"
'                            Query &= " , AddInfo = replace(concat(AddInfo,';CORECALCLATLON;'),';;',';')"
'                            Query &= " , UpdTime = now(), UpdUser = 'STI_CORECALCLATLON_BKO'"
'                            Query &= " Where TrackNum = '" & TrackNum & "'"
'                            ObjSQL.ExecNonQuery(MCon, Query)

'                            ErrMsg &= " " & ReqLatitude & " . " & ReqLongitude

'                        Else
'                            ErrMsg &= " " & "Gagal convert LatLon"

'                        End If

'                    Else
'                        'ErrMsg &= " " & "Sudah ada LatLon"

'                    End If
'                Catch ex As Exception
'                    ErrMsg &= " " & "Error convert LatLon" & ex.Message

'                End Try


'                'Insert TransactionDeliveryInfo
'                Query = "Insert Into TransactionDeliveryInfo ("
'                Query &= " TrackNum, OrderNo"
'                Query &= " , OrderType, OrderTypeDetail, ExpressType"
'                Query &= " , FlgCOD, CODValue, CODPaymentCode, CODPaymentBiller"
'                Query &= " , FlgBulky, ReturnBulkyItems"
'                Query &= " , CODValueDraft, ReturnBulkyItemsDraft"
'                Query &= " , ExpectedDeliverMinTime, ExpectedDeliverMaxTime"
'                Query &= " , ExpectedDeliverMinTimeDraft, ExpectedDeliverMaxTimeDraft"
'                Query &= " , OrderPaidTime, ReceiptPrintTime"
'                Query &= " , TimeZone"
'                Query &= " , BasePIN, BasePINDraft"
'                Query &= " , PINPickup, PINPickupDraft"
'                Query &= " , PINCancel, PINCancelDraft"
'                Query &= " , PINKeep, PINKeepDraft"
'                Query &= " , PINReturn, PINReturnDraft"
'                Query &= " , AddTime, AddUser"
'                Query &= " ) Select"

'                Query &= " '" & TrackNum & "', '" & OrderNo & "'"
'                Query &= " , 'PAKET' as OrderType, 'INDOPAKET' as OrderTypeDetail, 'SAMEDAY' as ExpressType"
'                Query &= " , '0' as FlgCOD, '0' as CODValue, '' as CODPaymentCode, '' as CODPaymentBiller"
'                Query &= " , '0' as FlgBulky, '' as ReturnBulkyItems"
'                Query &= " , '0' as CODValueDraft, '' as ReturnBulkyItemsDraft"
'                Query &= " , '" & ExpectedDeliverMinTime & "', '" & ExpectedDeliverMaxTime & "'"
'                Query &= " , '" & ExpectedDeliverMinTime & "', '" & ExpectedDeliverMaxTime & "'"
'                Query &= " , '" & SendRequestTime & "', '" & SendRequestTime & "'"
'                Query &= " , '" & DeliveryInfoTimeZone & "'"
'                Query &= " , '" & DeliveryInfoPINBase & "', '" & DeliveryInfoPINBase & "'"
'                Query &= " , '" & DeliveryInfoPINPickup & "', '" & DeliveryInfoPINPickup & "'"
'                Query &= " , '" & DeliveryInfoPINCancel & "', '" & DeliveryInfoPINCancel & "'"
'                Query &= " , '" & DeliveryInfoPINKeep & "', '" & DeliveryInfoPINKeep & "'"
'                Query &= " , '" & DeliveryInfoPINReturn & "', '" & DeliveryInfoPINReturn & "'"
'                Query &= " , now(), '" & User & "'"

'                If ObjSQL.ExecNonQuery(MCon, Query) = False Then
'                    ErrMsg = "Gagal Insert TransactionDeliveryInfo"
'                    GoTo SkipNext
'                End If


'                'Insert AutoOrderThirdPartyTransaction
'                Query = "Insert Into AutoOrderThirdPartyTransaction ("
'                Query &= " Account, ThirdParty, ServiceType"
'                Query &= " , TrackNum, OrderNo, dTrackNum"
'                Query &= " , Status1, Status2, SendRequestTo3PL"
'                Query &= " , AddTime, AddUser"
'                Query &= " ) Select"

'                Query &= " '" & Account & "', " & CustomAccount3rdPartyDeliveryMan & " as ThirdParty, 'SAMEDAY' as ServiceType"
'                Query &= " , '" & TrackNum & "', '" & OrderNo & "'"
'                Query &= " , cast(concat('" & TrackNum & "','01') as char) as dTrackNum"
'                Query &= " , '0', '0', '" & SendRequestTime & "'"
'                Query &= " , now(), '" & User & "'"

'                If ObjSQL.ExecNonQuery(MCon, Query) = False Then
'                    ErrMsg = "Gagal Insert AutoOrderThirdPartyTransaction"
'                    GoTo SkipNext
'                End If

'                '=== mirip dengan proses STI awb to-Door

'SkipNext:

'                dtTrackNumAll.Rows(i).Item("Status") = ErrMsg
'                dtTrackNumAll.AcceptChanges()

'                e.ResponseLog(MCon, AppName, AppVersion, Method, User, "SELESAI " & ErrMsg, LogKeyword & " " & TrackNum)

'            Next



'            Result(0) = "0"
'            Result(1) = ""
'            Result(2) = WService.ConvertDatatableToStringWithBarcode(dtTrackNumAll)

'Skip:

'            e.ResponseLog(MCon, AppName, AppVersion, Method, User, "SELESAI", LogKeyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, AppName, AppVersion, Method, "", ex, "", LogKeyword)

'            Result(1) = ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function

'    <WebMethod()>
'    Public Function ProcessTrackNumCoLatLon(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object() = CreateResult()

'        Dim Method As String = "ProcessTrackNumCoLatLon"

'        Dim MCon As MySqlConnection

'        Dim Query As String = ""

'        Dim LogKeyword As String = ""

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString
'            Dim TrackNumList As String = Param(2).ToString.Trim.ToUpper

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If TrackNumList = "" Then
'                Result(1) = "Input TrackNum"
'                GoTo Skip
'            End If

'            LogKeyword = User & " " & Format(Date.Now, "yyMMddHHmmss")
'            LogKeyword = WService.StrSQLRemoveChar(LogKeyword)

'            e.RequestLog(MCon, AppName, AppVersion, Method, User, "MULAI", LogKeyword)

'            Dim TrackNumAll As String() = TrackNumList.Split("|")

'            Dim dtTrackNumAll As New DataTable
'            dtTrackNumAll.Columns.Add("TrackNum")
'            dtTrackNumAll.Columns.Add("Status")

'            For i As Integer = 0 To TrackNumAll.Length - 1
'                If TrackNumAll(i) <> "" Then
'                    dtTrackNumAll.Rows.Add()
'                    dtTrackNumAll.Rows(dtTrackNumAll.Rows.Count - 1).Item("TrackNum") = TrackNumAll(i)
'                    dtTrackNumAll.Rows(dtTrackNumAll.Rows.Count - 1).Item("Status") = ""
'                    dtTrackNumAll.AcceptChanges()
'                End If
'            Next


'            If dtTrackNumAll.Rows.Count < 1 Then
'                Result(1) = "Input TrackNum "
'                GoTo Skip
'            End If

'            Dim MaxTrackNum As Integer = 50
'            If dtTrackNumAll.Rows.Count > MaxTrackNum Then
'                Result(1) = "Max. TrackNum " & MaxTrackNum
'                GoTo Skip
'            End If


'            Dim ObjSQL As New ClsSQL

'            For i As Integer = 0 To dtTrackNumAll.Rows.Count - 1

'                Dim ErrMsg As String = ""

'                Dim TrackNum As String = dtTrackNumAll.Rows(i).Item("TrackNum").ToString

'                e.RequestLog(MCon, AppName, AppVersion, Method, User, "MULAI", LogKeyword & " " & TrackNum)

'                Query = "Select t.TrackNum"
'                Query &= " , t.CoAddress, t.CoPostalCode"
'                Query &= " , t.CoLatitude, t.CoLongitude"
'                Query &= " From `Transaction` t"
'                Query &= " Where t.TrackNum = '" & TrackNum & "'"

'                Dim dtTransaction As DataTable = ObjSQL.ExecDatatable(MCon, Query)
'                If dtTransaction Is Nothing Then
'                    ErrMsg = "Gagal query"
'                    GoTo SkipNext
'                End If

'                If dtTransaction.Rows.Count < 1 Then
'                    ErrMsg = "Tidak ditemukan"
'                    GoTo SkipNext
'                End If

'                Try
'                    If dtTransaction.Rows(0).Item("CoLatitude").ToString = "" _
'                    Or dtTransaction.Rows(0).Item("CoLongitude").ToString = "" Then

'                        Dim dtLatLon As New DataTable
'                        dtLatLon.Columns.Add("REQADDRESS")
'                        dtLatLon.Columns.Add("REQPOSTALCODE")
'                        dtLatLon.Columns.Add("REQLATITUDE")
'                        dtLatLon.Columns.Add("REQLONGITUDE")

'                        Dim dr As DataRow
'                        dr = dtLatLon.NewRow

'                        dr.Item("REQADDRESS") = dtTransaction.Rows(0).Item("CoAddress").ToString
'                        dr.Item("REQPOSTALCODE") = dtTransaction.Rows(0).Item("CoPostalCode").ToString
'                        dr.Item("REQLATITUDE") = ""
'                        dr.Item("REQLONGITUDE") = ""

'                        dtLatLon.Rows.Add(dr)

'                        Dim ObjFungsi As New ClsFungsi
'                        Dim LatLonErrorMessage As String = ""
'                        Dim dtResult As DataTable = ObjFungsi.GoogleGeocodingLocation(dtLatLon, LatLonErrorMessage)

'                        Dim ReqLatitude As String = dtResult.Rows(0).Item("REQLATITUDE")
'                        Dim ReqLongitude As String = dtResult.Rows(0).Item("REQLONGITUDE")

'                        If LatLonErrorMessage <> "" Then
'                            ErrMsg = LatLonErrorMessage

'                        ElseIf ReqLatitude <> "" And ReqLongitude <> "" Then
'                            Query = "Update `Transaction`"
'                            Query &= " Set CoLatitude = '" & ReqLatitude & "', CoLongitude = '" & ReqLongitude & "'"
'                            Query &= " , AddInfo = replace(concat(AddInfo,';CORECALCLATLON;'),';;',';')"
'                            Query &= " , UpdTime = now(), UpdUser = 'STI_CORECALCLATLON_BKO'"
'                            Query &= " Where TrackNum = '" & TrackNum & "'"
'                            ObjSQL.ExecNonQuery(MCon, Query)

'                            ErrMsg = ReqLatitude & " . " & ReqLongitude

'                        Else
'                            ErrMsg = "Gagal convert"

'                        End If

'                    Else
'                        ErrMsg = "Sudah ada LatLon"

'                    End If

'                Catch ex As Exception
'                    ErrMsg = ex.Message

'                End Try

'SkipNext:

'                dtTrackNumAll.Rows(i).Item("Status") = ErrMsg
'                dtTrackNumAll.AcceptChanges()

'                e.ResponseLog(MCon, AppName, AppVersion, Method, User, "SELESAI " & ErrMsg, LogKeyword & " " & TrackNum)

'            Next


'            Result(0) = "0"
'            Result(1) = ""
'            Result(2) = "" 'WService.ConvertDatatableToStringWithBarcode(dtTrackNumAll)

'Skip:

'            e.ResponseLog(MCon, AppName, AppVersion, Method, User, "SELESAI", LogKeyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, "", "", Method, User, ex, "", LogKeyword)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try
'        End Try

'        Return Result

'    End Function


'    <WebMethod()>
'    Public Function SyncTokoData(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object() = CreateResult()

'        Dim Method As String = "SyncTokoData"

'        Dim MCon As MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, User, "MULAI")

'            Dim ObjSQL As New ClsSQL

'            'Dim JenisToko As String = "'MOTHER STORE','SS MOTHER STORE','STOCKPOINT GUDANG','SPECIAL STORE','POINT COFFEE VIRTUAL'"
'            'Dim JenisToko As String = "'MOTHER STORE','MS EXPRESS','SS MOTHER STORE','SS MS EXPRESS','SPECIAL STORE','STOCKPOINT GUDANG'"

'            Dim JenisToko As String = "'MOTHER STORE','SS MOTHER STORE','STOCKPOINT GUDANG'"

'            Query = "Drop Table If Exists IndomaretStore_DeliveryMan_Backup"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Create Table IndomaretStore_DeliveryMan_Backup Like IndomaretStore_DeliveryMan"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Insert Into IndomaretStore_DeliveryMan_Backup Select * From IndomaretStore_DeliveryMan"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            'initial data perlu dilaporkan ke sistem deliveryman
'            Query = "Insert Into IndomaretStore_DeliveryMan ("
'            Query &= " Code, `Status`"
'            Query &= " , DcCode, OpHourOpen, OpHourClose, InactiveDate" '-- kolom yang mungkin mengalami perubahan; perlu dilaporkan lagi ke sistem backoffice deliman
'            Query &= " , AddTime, AddUser"
'            Query &= " )"

'            Query &= " Select i.KodeToko as Code, '8' as `Status`"
'            Query &= " , i.KodeGudang, i.Tok_Jam_Buka, i.Tok_Jam_Tutup, i.Tok_Tgl_Tutup"
'            Query &= " , now(), 'INIT'"
'            Query &= " From indomaretstore_temp i" '-- table original dari oracle idm (sd4, file ORG)
'            'Query &= " Left Join indomaretstore_deliveryman d on (d.Code = i.KodeToko)"

'            Query &= " Where True"
'            'Query &= " And d.Code is null" '-- belum pernah diproses
'            Query &= " And i.JenisTokoEco in (" & JenisToko & ")" '-- kriteria data toko yang perlu diproses

'            Query &= " On Duplicate Key Update"
'            Query &= " DcCode = i.KodeGudang, OpHourOpen = i.Tok_Jam_Buka, OpHourClose = i.Tok_Jam_Tutup, InactiveDate = i.Tok_Tgl_Tutup"

'            ObjSQL.ExecNonQuery(MCon, Query)

'            'catat latitude longitude
'            Query = "Update indomaretstore_deliveryman d"
'            Query &= " Inner Join indomaretstore_for_report i on (i.Code = d.Code)"
'            Query &= " Set d.Latitude = i.Latitude, d.Longitude = i.Longitude"
'            ObjSQL.ExecNonQuery(MCon, Query)


'            'pastikan aktif BKO di IPP
'            Query = "insert ignore into IndomaretStore_BkoIpp ("
'            Query &= " Code, ActiveDate, AddTime, AddUser"
'            Query &= " ) Select"
'            Query &= " d.Code, curdate(), now(), 'SyncDataDeliman'"
'            Query &= " From IndomaretStore_DeliveryMan d"
'            Query &= " Left Join Indomaretstore_BkoIpp b on (b.Code = d.Code)"
'            Query &= " Where b.Code is null"
'            ObjSQL.ExecNonQuery(MCon, Query)


'            Query = "Update IndomaretStore_DeliveryMan d"
'            Query &= " Inner Join MstLogin l on (l.`Type` = 'DSH' And l.Code = '1000319' and l.`User` = concat('MS',d.Code))" '-- perubahan pada dashboard KlikIndomaretGrab
'            Query &= " Set l.UpdTime = now(), l.UpdUser = 'INIT IdmStoreDeliveryMan'"
'            Query &= " , l.UserGroup = 'MOTHERSTOREDLM'" '-- perubahan menu dashboard yang ditampilkan
'            Query &= " , l.AddInfo = replace(l.AddInfo,'SENDTELEGRAM=','SENDXYZTELEGRAM=')" '-- tidak lagi send telegram ke grup third party instant
'            Query &= " Where d.`Status` = '8'"
'            ObjSQL.ExecNonQuery(MCon, Query)


'            '-- insert user baru untuk dashboard KlikIndomaretGrab
'            Query = "Insert Ignore Into MstLogin ("
'            Query &= " `User`, `Password`, `Name`, `Type`, Code, UserGroup, AddInfo, ActiveDate, AddTime, AddUser"
'            Query &= " ) Select"
'            Query &= " concat('MS',Code) as `User`, concat('MS',Code) as `Password`, concat('Motherstore ',Code) as `Name`"
'            Query &= " , 'DSH' as `Type`, '1000319' as Code, 'MOTHERSTOREDLM' as UserGroup"
'            Query &= " , concat('MSTORE=',Code,';AUTOORDERTHIRDPARTY=Y;TRACKUSERAWB=Y;') as AddInfo"
'            Query &= " , curdate() as ActiveDate, now() as AddTime, 'INIT IdmStoreDeliveryMan' as AddUser"
'            Query &= " From IndomaretStore_DeliveryMan"
'            Query &= " Where `Status` = '8'"
'            ObjSQL.ExecNonQuery(MCon, Query)
'            'sudah ditambahkan trigger di database, agar ketika insert user baru yang user = pass, maka akan password akan di-random sesuai ketentuan It Security 240910

'            '-- username dan password sudah di-random, jadi perlu di-info ke SPV DMS via email
'            Query = " Select l.`User`, l.`Password`, l.`Name`"
'            Query &= " From IndomaretStore_DeliveryMan d"
'            Query &= " From MstLogin l on (l.`User` = concat('MS',d.Code) and d.`Type`='DSH' and d.Code = '1000319')"
'            Query &= " Where d.`Status` = '8'"



'            'tentukan data toko mana saja yang saat ini aktif untuk dilaporkan
'            Query = "Update indomaretstore_deliveryman d"
'            Query &= " Inner Join indomaretstore_for_report i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = cast(case when curdate() between i.ActiveDate and ifnull(i.InActiveDate,curdate()) then '0' else '9' end as char)"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'ISACTIVE'"
'            Query &= " Where d.`Status` = '8'"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            'tentukan data toko mana saja yang ada perubahan informasi
'            Query = "Update IndomaretStore_DeliveryMan d"
'            Query &= " Inner Join IndomaretStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_DCCODE'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.DcCode,'') <> ifnull(d.DcCode,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Update IndomaretStore_DeliveryMan d"
'            Query &= " Inner Join IndomaretStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_HOUROPEN'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.OpHourOpen,'') <> ifnull(d.OpHourOpen,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Update IndomaretStore_DeliveryMan d"
'            Query &= " Inner Join IndomaretStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_HOURCLOSE'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.OpHourClose,'') <> ifnull(d.OpHourClose,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Update IndomaretStore_DeliveryMan d"
'            Query &= " Inner Join IndomaretStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_INACTIVE'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.InactiveDate,'') <> ifnull(d.InactiveDate,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Update IndomaretStore_DeliveryMan d"
'            Query &= " Inner Join IndomaretStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_LAT'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.Latitude,'') <> '0' And ifnull(i.Latitude,'') <> ifnull(d.Latitude,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Update IndomaretStore_DeliveryMan d"
'            Query &= " Inner Join IndomaretStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_LON'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.Longitude,'') <> '0' And ifnull(i.Longitude,'') <> ifnull(d.Longitude,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)


'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = ""

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, User, "SELESAI " & Result(0).ToString)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog("", "", Method, User, ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try


'        Return Result

'    End Function

'    <WebMethod()>
'    Public Function SyncTokoSend(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object() = CreateResult()

'        Dim Method As String = "SyncTokoSend"

'        Dim MCon As MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim LogMessage As String = ""

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, User, "MULAI")

'            Dim ObjSQL As New ClsSQL

'            'Dim TbName As String = "TempSyncTokoDeliveryMan_" & Format(Date.Now, "yyMMddHHmmss")
'            Dim TbName As String = "TempSyncTokoDeliveryMan"

'            ObjSQL.ExecNonQuery(MCon, "Drop Temporary Table If Exists " & TbName)

'            'create temporary table
'            Query = "Create Temporary Table " & TbName
'            Query &= " Select d.Code, ifnull(d.UpdTime,d.AddTime) as AddTime"
'            'Query &= " , cast('00:00' as char) as JamBuka, cast('00:00' as char) as JamTutup" '-- informasi tambahan yang akan diperlukan
'            Query &= " , cast(d.OpHourOpen as char) as JamBuka, cast(d.OpHourClose as char) as JamTutup"
'            Query &= " From IndomaretStore_DeliveryMan d"
'            Query &= " Where d.`Status` = '0'" '-- yang perlu diproses
'            Query &= " Order By d.AddTime, d.UpdTime"
'            'Query &= " Limit 20"
'            'Query &= " Limit 100"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            ObjSQL.ExecNonQuery(MCon, "Alter Table " & TbName & " Add Primary Key (`Code`)")


'            ''update temporary table untuk informasi tambahan yang diperlukan
'            'Query = "Update " & TbName & " d"
'            'Query &= " Inner Join indomaretstore_temp i on (i.KodeToko = d.Code)"
'            'Query &= " Set d.JamBuka = i.Tok_Jam_Buka, d.JamTutup = i.Tok_Jam_Tutup"
'            'ObjSQL.ExecNonQuery(MCon, Query)


'            'query informasi yang perlu dilaporkan
'            Query = "Select d.Code"

'            Query &= " , i.Name, i.Phone"
'            Query &= " , i.Address, i.PostalCode"
'            Query &= " , i.City as Kota, i.Kecamatan, i.Kelurahan"
'            Query &= " , i.Latitude, i.Longitude"
'            Query &= " , i.DcCode, i.DcName"
'            Query &= " , cast(curdate() between i.ActiveDate and ifnull(i.InActiveDate, curdate()) as char) as IsActive"
'            Query &= " , cast(concat(d.JamBuka,':00','+07:00') as char) as JamBuka, cast(concat(d.JamTutup,':00','+07:00') as char) as JamTutup"
'            Query &= " , cast(date_format(d.AddTime,'%Y-%m-%dT%H:%i:%s+07:00') as char) as AddTime"

'            Query &= " From `" & TbName & "` d"
'            Query &= " Inner Join indomaretstore_for_report i on (i.Code = d.Code)" '-- informasi utama dari table indomaretstore

'            Query &= " Where True"

'            Dim dtToko As DataTable = ObjSQL.ExecDatatable(MCon, Query)

'            If dtToko.Rows.Count > 0 Then

'                Dim ObjFungsi As New ClsFungsi

'                'init oauth
'                Dim Access_Token As String = ""

'                Dim tParam(1) As Object
'                tParam(0) = Param(0)
'                tParam(1) = Param(1)

'                Dim tResult As Object() = GetOAuth(wsAppName, wsAppVersion, tParam)
'                If tResult(0).ToString = "0" Then
'                    Access_Token = tResult(2).ToString
'                Else
'                    Result(1) = tResult(1).ToString
'                    GoTo Skip
'                End If

'                ReDim CustomHeaders(0)
'                CustomHeaders(0) = "X-API-Key|" & Access_Token

'                Dim ParameterHeaders As String = ""
'                For h As Integer = 0 To CustomHeaders.Length - 1
'                    If ParameterHeaders <> "" Then
'                        ParameterHeaders &= ","
'                    End If
'                    ParameterHeaders &= CustomHeaders(h)
'                Next


'                Dim SendOK As Integer = 0
'                Dim SendNOK As Integer = 0

'                For i As Integer = 0 To dtToko.Rows.Count - 1

'                    Dim ObjSyncToko As New cSyncTokoDeliveryMan
'                    ObjSyncToko.storeCode = dtToko.Rows(i).Item("Code").ToString
'                    ObjSyncToko.storeName = dtToko.Rows(i).Item("Name").ToString
'                    ObjSyncToko.contact = dtToko.Rows(i).Item("Phone").ToString
'                    ObjSyncToko.address = dtToko.Rows(i).Item("Address").ToString
'                    ObjSyncToko.postalCode = dtToko.Rows(i).Item("PostalCode").ToString
'                    ObjSyncToko.city = dtToko.Rows(i).Item("Kota").ToString
'                    ObjSyncToko.subDistrict = dtToko.Rows(i).Item("Kecamatan").ToString
'                    ObjSyncToko.urban = dtToko.Rows(i).Item("Kelurahan").ToString
'                    ObjSyncToko.latitude = CDbl(dtToko.Rows(i).Item("Latitude"))
'                    ObjSyncToko.longitude = CDbl(dtToko.Rows(i).Item("Longitude"))
'                    ObjSyncToko.openingHour = dtToko.Rows(i).Item("JamBuka").ToString
'                    ObjSyncToko.closingHour = dtToko.Rows(i).Item("JamTutup").ToString
'                    ObjSyncToko.lastModifiedDate = dtToko.Rows(i).Item("AddTime").ToString
'                    ObjSyncToko.dcCode = dtToko.Rows(i).Item("DcCode").ToString
'                    ObjSyncToko.dcName = dtToko.Rows(i).Item("DcName").ToString
'                    ObjSyncToko.active = (dtToko.Rows(i).Item("IsActive").ToString = "1")
'                    ObjSyncToko.channel = "INDOMARET"

'                    Dim Parameter As String = JsonConvert.SerializeObject(ObjSyncToko)

'                    Dim AlreadyResetToken As Boolean = False
'                    Dim MustResetToken As Boolean = False

'UlangGetOAuth:

'                    If MustResetToken Then

'                        Access_Token = ""

'                        ReDim tParam(2)
'                        tParam(0) = Param(0)
'                        tParam(1) = Param(1)
'                        tParam(2) = "1"
'                        AlreadyResetToken = True

'                        tResult = GetOAuth(wsAppName, wsAppVersion, tParam)
'                        If tResult(0).ToString = "0" Then
'                            Access_Token = tResult(2).ToString
'                        Else
'                            Result(1) = tResult(1).ToString
'                            GoTo Skip
'                        End If

'                        ReDim CustomHeaders(0)
'                        CustomHeaders(0) = "X-API-Key|" & Access_Token

'                        ParameterHeaders = ""
'                        For h As Integer = 0 To CustomHeaders.Length - 1
'                            If ParameterHeaders <> "" Then
'                                ParameterHeaders &= ","
'                            End If
'                            ParameterHeaders &= CustomHeaders(h)
'                        Next

'                    End If 'dari If MustResetToken

'                    e.APIRequestLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "Header:" & ParameterHeaders & "-" & "Body:" & Parameter, mURLSyncToko, ObjSyncToko.storeCode)

'                    Dim Response As String = ""
'                    Response = ObjFungsi.SendHTTP("", "", "", mURLSyncToko, Parameter, "", Encoding.Default, "60", "", mContentType, True, CustomHeaders)
'                    Response = ("" & Response).Trim

'                    If Response <> "" Then

'                        e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, Response, mURLSyncToko, ObjSyncToko.storeCode)

'                        If Response.ToUpper.Contains("00") And Response.ToUpper.Contains("SUCCESS") Then
'                            Query = "Update IndomaretStore_DeliveryMan"
'                            Query &= " Set `Status` = '1'"
'                            Query &= " , UpdTime = now(), UpdUser = 'SEND'"
'                            Query &= " Where Code = '" & ObjSyncToko.storeCode & "'"
'                            ObjSQL.ExecNonQuery(MCon, Query)

'                            SendOK += 1
'                        Else
'                            Query = "Update IndomaretStore_DeliveryMan"
'                            Query &= " Set `Status` = '7'"
'                            Query &= " , UpdTime = now(), UpdUser = '" & Left("ERR " & WService.StrSQLRemoveChar(Response.Trim), 45) & "'"
'                            Query &= " Where Code = '" & ObjSyncToko.storeCode & "'"
'                            ObjSQL.ExecNonQuery(MCon, Query)

'                            SendNOK += 1
'                        End If

'                    Else

'                        e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "No Response", mURLSyncToko, ObjSyncToko.storeCode)

'                        SendNOK += 1

'                    End If 'dari If Response <> ""

'                Next

'                LogMessage &= SendOK & " " & SendNOK

'            Else

'                LogMessage &= "TIDAK ADA YANG DIPROSES"

'            End If 'dari If dtToko.Rows.Count > 0


'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = ""

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, User, "SELESAI " & Result(0).ToString & " " & LogMessage)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog("", "", Method, User, ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try


'        Return Result

'    End Function

'    Public Class cSyncTokoDeliveryMan
'        Public storeCode As String = ""
'        Public storeName As String = ""
'        Public address As String = ""
'        Public contact As String = ""
'        Public latitude As Double = 0
'        Public longitude As Double = 0
'        Public postalCode As String = ""
'        Public city As String = ""
'        Public subDistrict As String = ""
'        Public urban As String = ""
'        Public openingHour As String = ""
'        Public closingHour As String = ""
'        Public active As Boolean = True
'        Public lastModifiedDate As String = ""
'        Public dcCode As String = ""
'        Public dcName As String = ""
'        Public channel As String = ""
'    End Class


'    <WebMethod()>
'    Public Function SyncTokoDataIgr(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object() = CreateResult()

'        Dim Method As String = "SyncTokoDataIgr"

'        Dim MCon As MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, User, "MULAI")

'            Dim ObjSQL As New ClsSQL


'            Query = "Drop Table If Exists IndogrosirStore_DeliveryMan_Backup"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Create Table IndogrosirStore_DeliveryMan_Backup Like IndogrosirStore_DeliveryMan"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Insert Into IndogrosirStore_DeliveryMan_Backup Select * From IndogrosirStore_DeliveryMan"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            'initial data perlu dilaporkan ke sistem deliveryman
'            Query = "Insert Into IndogrosirStore_DeliveryMan ("
'            Query &= " Code, `Status`"
'            Query &= " , DcCode, OpHourOpen, OpHourClose, InactiveDate" '-- kolom yang mungkin mengalami perubahan; perlu dilaporkan lagi ke sistem backoffice deliman
'            Query &= " , AddTime, AddUser"
'            Query &= " )"

'            Query &= " Select i.Code, '8' as `Status`"
'            Query &= " , i.NumCode, '08:00' as Tok_Jam_Buka, '20:00' as Tok_Jam_Tutup, ifnull(i.InActiveDate,'') as Tok_Tgl_Tutup"
'            Query &= " , now(), 'INIT'"
'            Query &= " From indogrosirstore i" '-- table original

'            Query &= " Where i.FlgNewTrx = 'Y'"

'            Query &= " On Duplicate Key Update"
'            Query &= " DcCode = i.NumCode, OpHourOpen = '08:00', OpHourClose = '20:00', InactiveDate = ifnull(i.InActiveDate,'')"

'            ObjSQL.ExecNonQuery(MCon, Query)

'            'catat latitude longitude
'            Query = "Update indogrosirstore_deliveryman d"
'            Query &= " Inner Join indogrosirstore i on (i.Code = d.Code)"
'            Query &= " Set d.Latitude = i.Latitude, d.Longitude = i.Longitude"
'            ObjSQL.ExecNonQuery(MCon, Query)


'            'tentukan data toko mana saja yang saat ini aktif untuk dilaporkan
'            Query = "Update indogrosirstore_deliveryman d"
'            Query &= " Inner Join indogrosirstore i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = cast(case when curdate() between i.ActiveDate and ifnull(i.InActiveDate,curdate()) then '0' else '9' end as char)"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'ISACTIVE'"
'            Query &= " Where d.`Status` = '8'"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            'tentukan data toko mana saja yang ada perubahan informasi
'            Query = "Update IndogrosirStore_DeliveryMan d"
'            Query &= " Inner Join IndogrosirStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_DCCODE'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.DcCode,'') <> ifnull(d.DcCode,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Update IndogrosirStore_DeliveryMan d"
'            Query &= " Inner Join IndogrosirStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_HOUROPEN'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.OpHourOpen,'') <> ifnull(d.OpHourOpen,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Update IndogrosirStore_DeliveryMan d"
'            Query &= " Inner Join IndogrosirStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_HOURCLOSE'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.OpHourClose,'') <> ifnull(d.OpHourClose,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Update IndogrosirStore_DeliveryMan d"
'            Query &= " Inner Join IndogrosirStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_INACTIVE'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.InactiveDate,'') <> ifnull(d.InactiveDate,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Update IndogrosirStore_DeliveryMan d"
'            Query &= " Inner Join IndogrosirStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_LAT'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.Latitude,'') <> '0' And ifnull(i.Latitude,'') <> ifnull(d.Latitude,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)

'            Query = "Update IndogrosirStore_DeliveryMan d"
'            Query &= " Inner Join IndogrosirStore_DeliveryMan_Backup i on (i.Code = d.Code)"
'            Query &= " Set d.`Status` = 0"
'            Query &= " , d.UpdTime = now(), d.UpdUser = 'UPD_LON'"
'            Query &= " Where d.`Status` = '1'"
'            Query &= " And ifnull(i.Longitude,'') <> '0' And ifnull(i.Longitude,'') <> ifnull(d.Longitude,'')"
'            ObjSQL.ExecNonQuery(MCon, Query)


'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = ""

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, User, "SELESAI " & Result(0).ToString)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog("", "", Method, User, ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try


'        Return Result

'    End Function

'    <WebMethod()>
'    Public Function SyncTokoSendIgr(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object() = CreateResult()

'        Dim Method As String = "SyncTokoSendIgr"

'        Dim MCon As MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim LogMessage As String = ""

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, User, "MULAI")

'            Dim ObjSQL As New ClsSQL

'            'Dim TbName As String = "TempSyncTokoDeliveryManIgr_" & Format(Date.Now, "yyMMddHHmmss")
'            Dim TbName As String = "TempSyncTokoDeliveryManIgr"

'            ObjSQL.ExecNonQuery(MCon, "Drop Temporary Table If Exists " & TbName)

'            'create temporary table
'            Query = "Create Temporary Table " & TbName
'            Query &= " Select d.Code, ifnull(d.UpdTime,d.AddTime) as AddTime"
'            Query &= " , cast(d.OpHourOpen as char) as JamBuka, cast(d.OpHourClose as char) as JamTutup"
'            Query &= " From IndogrosirStore_DeliveryMan d"
'            Query &= " Where d.`Status` = '0'" '-- yang perlu diproses
'            Query &= " Order By d.AddTime, d.UpdTime"

'            ObjSQL.ExecNonQuery(MCon, Query)

'            ObjSQL.ExecNonQuery(MCon, "Alter Table " & TbName & " Add Primary Key (`Code`)")


'            'query informasi yang perlu dilaporkan
'            Query = "Select d.Code, cast(concat(i.PrefixDmsCode,d.Code) as char) as DmsCode"

'            Query &= " , i.Name, i.Phone"
'            Query &= " , i.Address, i.PostalCode"
'            Query &= " , i.City as Kota, i.Kecamatan, i.Kelurahan"
'            Query &= " , i.Latitude, i.Longitude"
'            Query &= " , i.NumCode as DcCode, i.Code as DcName"
'            Query &= " , cast(curdate() between i.ActiveDate and ifnull(i.InActiveDate, curdate()) as char) as IsActive"
'            Query &= " , cast(concat(d.JamBuka,':00','+07:00') as char) as JamBuka, cast(concat(d.JamTutup,':00','+07:00') as char) as JamTutup"
'            Query &= " , cast(date_format(d.AddTime,'%Y-%m-%dT%H:%i:%s+07:00') as char) as AddTime"

'            Query &= " From `" & TbName & "` d"
'            Query &= " Inner Join indogrosirstore i on (i.Code = d.Code)" '-- informasi utama dari table indogrosir

'            Query &= " Where True"

'            Dim dtToko As DataTable = ObjSQL.ExecDatatable(MCon, Query)

'            If dtToko.Rows.Count > 0 Then

'                Dim ObjFungsi As New ClsFungsi

'                'init oauth
'                Dim Access_Token As String = ""

'                Dim tParam(1) As Object
'                tParam(0) = Param(0)
'                tParam(1) = Param(1)

'                Dim tResult As Object() = GetOAuth(wsAppName, wsAppVersion, tParam)
'                If tResult(0).ToString = "0" Then
'                    Access_Token = tResult(2).ToString
'                Else
'                    Result(1) = tResult(1).ToString
'                    GoTo Skip
'                End If

'                ReDim CustomHeaders(0)
'                CustomHeaders(0) = "X-API-Key|" & Access_Token

'                Dim ParameterHeaders As String = ""
'                For h As Integer = 0 To CustomHeaders.Length - 1
'                    If ParameterHeaders <> "" Then
'                        ParameterHeaders &= ","
'                    End If
'                    ParameterHeaders &= CustomHeaders(h)
'                Next


'                Dim SendOK As Integer = 0
'                Dim SendNOK As Integer = 0

'                For i As Integer = 0 To dtToko.Rows.Count - 1

'                    Dim StoreCode As String = dtToko.Rows(i).Item("Code").ToString

'                    Dim ObjSyncToko As New cSyncTokoDeliveryMan
'                    ObjSyncToko.storeCode = dtToko.Rows(i).Item("DmsCode").ToString
'                    ObjSyncToko.storeName = dtToko.Rows(i).Item("Name").ToString
'                    ObjSyncToko.contact = dtToko.Rows(i).Item("Phone").ToString
'                    ObjSyncToko.address = dtToko.Rows(i).Item("Address").ToString
'                    ObjSyncToko.postalCode = dtToko.Rows(i).Item("PostalCode").ToString
'                    ObjSyncToko.city = dtToko.Rows(i).Item("Kota").ToString
'                    ObjSyncToko.subDistrict = dtToko.Rows(i).Item("Kecamatan").ToString
'                    ObjSyncToko.urban = dtToko.Rows(i).Item("Kelurahan").ToString
'                    ObjSyncToko.latitude = CDbl(dtToko.Rows(i).Item("Latitude"))
'                    ObjSyncToko.longitude = CDbl(dtToko.Rows(i).Item("Longitude"))
'                    ObjSyncToko.openingHour = dtToko.Rows(i).Item("JamBuka").ToString
'                    ObjSyncToko.closingHour = dtToko.Rows(i).Item("JamTutup").ToString
'                    ObjSyncToko.lastModifiedDate = dtToko.Rows(i).Item("AddTime").ToString
'                    ObjSyncToko.dcCode = dtToko.Rows(i).Item("DcCode").ToString
'                    ObjSyncToko.dcName = dtToko.Rows(i).Item("DcName").ToString
'                    ObjSyncToko.active = (dtToko.Rows(i).Item("IsActive").ToString = "1")
'                    ObjSyncToko.channel = "INDOGROSIR"

'                    Dim Parameter As String = JsonConvert.SerializeObject(ObjSyncToko)

'                    Dim AlreadyResetToken As Boolean = False
'                    Dim MustResetToken As Boolean = False

'UlangGetOAuth:

'                    If MustResetToken Then

'                        Access_Token = ""

'                        ReDim tParam(2)
'                        tParam(0) = Param(0)
'                        tParam(1) = Param(1)
'                        tParam(2) = "1"
'                        AlreadyResetToken = True

'                        tResult = GetOAuth(wsAppName, wsAppVersion, tParam)
'                        If tResult(0).ToString = "0" Then
'                            Access_Token = tResult(2).ToString
'                        Else
'                            Result(1) = tResult(1).ToString
'                            GoTo Skip
'                        End If

'                        ReDim CustomHeaders(0)
'                        CustomHeaders(0) = "X-API-Key|" & Access_Token

'                        ParameterHeaders = ""
'                        For h As Integer = 0 To CustomHeaders.Length - 1
'                            If ParameterHeaders <> "" Then
'                                ParameterHeaders &= ","
'                            End If
'                            ParameterHeaders &= CustomHeaders(h)
'                        Next

'                    End If 'dari If MustResetToken

'                    e.APIRequestLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "Header:" & ParameterHeaders & "-" & "Body:" & Parameter, mURLSyncToko, ObjSyncToko.storeCode)

'                    Dim Response As String = ""
'                    Response = ObjFungsi.SendHTTP("", "", "", mURLSyncToko, Parameter, "", Encoding.Default, "60", "", mContentType, True, CustomHeaders)
'                    Response = ("" & Response).Trim

'                    If Response <> "" Then

'                        e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, Response, mURLSyncToko, ObjSyncToko.storeCode)

'                        If Response.ToUpper.Contains("00") And Response.ToUpper.Contains("SUCCESS") Then
'                            Query = "Update IndogrosirStore_DeliveryMan"
'                            Query &= " Set `Status` = '1'"
'                            Query &= " , UpdTime = now(), UpdUser = 'SEND'"
'                            Query &= " Where Code = '" & StoreCode & "'"
'                            ObjSQL.ExecNonQuery(MCon, Query)

'                            SendOK += 1
'                        Else
'                            Query = "Update IndogrosirStore_DeliveryMan"
'                            Query &= " Set `Status` = '7'"
'                            Query &= " , UpdTime = now(), UpdUser = '" & Left("ERR " & WService.StrSQLRemoveChar(Response.Trim), 45) & "'"
'                            Query &= " Where Code = '" & StoreCode & "'"
'                            ObjSQL.ExecNonQuery(MCon, Query)

'                            SendNOK += 1
'                        End If

'                    Else

'                        e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "No Response", mURLSyncToko, ObjSyncToko.storeCode)

'                        SendNOK += 1

'                    End If 'dari If Response <> ""

'                Next

'                LogMessage &= SendOK & " " & SendNOK

'            Else

'                LogMessage &= "TIDAK ADA YANG DIPROSES"

'            End If 'dari If dtToko.Rows.Count > 0


'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = ""

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, User, "SELESAI " & Result(0).ToString & " " & LogMessage)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog("", "", Method, User, ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try


'        Return Result

'    End Function


'    'Private wsSlave As New WServiceSlave.ServiceSlave


'    Public mPPIdmURLOAuth As String = ""
'    Public mPPIdmURLPaidStatus As String = ""
'    Public mPPIdmContentType As String = ""
'    Public wsUserPPIdm As String = ""
'    Public wsPassPPIdm As String = ""

'    Public Sub PaymentPointIndomaretInit()

'        If mPPIdmURLOAuth = "" Then
'            Try
'                mPPIdmURLOAuth = ("" & ConfigurationManager.AppSettings("PaymentPointIdmUrlOAuth")).Trim
'            Catch ex As Exception
'                mPPIdmURLOAuth = ""
'            End Try
'            If mPPIdmURLOAuth = "" Then
'                mPPIdmURLOAuth = "x"
'            End If
'        End If

'        If mPPIdmURLPaidStatus = "" Then
'            Try
'                mPPIdmURLPaidStatus = ("" & ConfigurationManager.AppSettings("PaymentPointIdmUrlPaidStatus")).Trim
'            Catch ex As Exception
'                mPPIdmURLPaidStatus = ""
'            End Try
'            If mPPIdmURLPaidStatus = "" Then
'                mPPIdmURLPaidStatus = "x"
'            End If
'        End If

'        If mPPIdmContentType = "" Then
'            mPPIdmContentType = "application/json"
'        End If

'        If wsUserPPIdm = "" Then
'            Try
'                wsUserPPIdm = ("" & ConfigurationManager.AppSettings("PaymentPointIdmWsUser")).Trim
'            Catch ex As Exception
'                wsUserPPIdm = ""
'            End Try
'            If wsUserPPIdm = "" Then
'                wsUserPPIdm = "x"
'            End If
'        End If

'        If wsPassPPIdm = "" Then
'            Try
'                wsPassPPIdm = ("" & ConfigurationManager.AppSettings("PaymentPointIdmWsPassword")).Trim
'            Catch ex As Exception
'                wsPassPPIdm = ""
'            End Try
'            If wsPassPPIdm = "" Then
'                wsPassPPIdm = "x"
'            End If
'        End If

'    End Sub

'    Public Class cPPIdmReqOAuth
'        Public Produk As String = "PAYMENTPOINT"
'        Public Action As String = "LOGIN"
'        Public Detail As New cPPIdmReqOAuthDetail
'    End Class

'    Public Class cPPIdmReqOAuthDetail
'        Public Username As String = ""
'        Public Password As String = ""
'    End Class

'    Public Class cPPIdmRspOAuth
'        Public RespCode As String = "" '0 = sukses
'        Public RespDesc As String = ""
'        Public Detail As New cPPIdmRspOAuthDetail
'    End Class

'    Public Class cPPIdmRspOAuthDetail
'        Public Key As String = "" 'token / oauth
'    End Class

'    <WebMethod()>
'    Public Function PaymentPointIndomaretGetOAuth(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object

'        Dim Method As String = "PaymentPointIndomaretGetOAuth"

'        Dim Result As Object() = CreateResult()
'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString

'            Dim MustReset As Boolean = False
'            Try
'                MustReset = (Param(2).ToString = "1")
'            Catch ex As Exception
'                MustReset = False
'            End Try


'            MCon = MasterMCon.Clone

'            Dim ObjFungsi As New ClsFungsi

'            Dim e As New ClsError

'            Dim ObjSQL As New ClsSQL

'            PaymentPointIndomaretInit()

'            Dim Access_Token As String = ""

'            Dim AccountPPIndomaret As String = "PPINDOMARET"

'            If MustReset = False Then
'                'get local oauth
'                Query = "Select OAuth From AutoOrderOAuth"
'                Query &= " Where `User` = '" & AccountPPIndomaret & "'"
'                Try
'                    Access_Token = ObjSQL.ExecScalar(MCon, Query).ToString
'                Catch ex As Exception
'                    Access_Token = ""
'                End Try
'            End If 'If MustReset = False

'            If Access_Token = "" Then

'                'get new oauth
'                Dim ObjReqOAuth As New cPPIdmReqOAuth

'                ObjReqOAuth.Detail.Username = wsUserPPIdm
'                ObjReqOAuth.Detail.Password = wsPassPPIdm

'                Dim Parameter As String = JsonConvert.SerializeObject(ObjReqOAuth)

'                e.APIRequestLog(MCon, wsAppName, wsAppVersion, Method, wsUser, Parameter, mPPIdmURLOAuth)

'                'Dim Response As String = ""
'                'Response = ObjFungsi.SendHTTP("", "", "", mPPIdmURLOAuth, Parameter, "", Encoding.Default, "60", "", mPPIdmContentType)
'                'Response = ("" & Response).Trim

'                Dim oParam(2) As Object
'                oParam(0) = Parameter
'                oParam(1) = mPPIdmURLOAuth
'                oParam(2) = mPPIdmContentType

'                wsSlave.Url = ("" & ConfigurationManager.AppSettings("WServiceSlave.ServiceSlave")).Trim

'                Dim Response As String = wsSlave.PaymentPointIndomaretGetOAuthV2(oParam)

'                Dim ObjRspOAuth As New cPPIdmRspOAuth

'                If Response <> "" Then
'                    e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsUser, Response, mPPIdmURLOAuth)
'                    Try
'                        ObjRspOAuth = JsonConvert.DeserializeObject(Of cPPIdmRspOAuth)(Response)
'                        Access_Token = ObjRspOAuth.Detail.Key
'                    Catch ex As Exception
'                        e.ErrorLog(wsAppName, wsAppVersion, Method, wsUser, ex, "")
'                        Access_Token = ""
'                    End Try
'                Else
'                    e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsUser, "No Response", mPPIdmURLOAuth)
'                End If 'dari If Response <> ""

'                If Access_Token.Trim = "" Then
'                    Result(1) = "Gagal GetOAuth"
'                    GoTo Skip
'                End If

'                'save local oauth
'                Query = "Insert Into AutoOrderOAuth ("
'                Query &= " `User`, OAuth, UpdTime, UpdUser"
'                Query &= " ) values ("
'                Query &= " '" & AccountPPIndomaret & "','" & Access_Token & "'"
'                Query &= " , now(), '" & wsUser & "'"
'                Query &= " ) on duplicate key update"
'                Query &= " OAuth = '" & Access_Token & "'"
'                Query &= " , UpdTime = now(), UpdUser = '" & wsUser & "'"
'                ObjSQL.ExecScalar(MCon, Query)

'            End If 'dari If Access_Token = ""


'            ReDim Result(2)
'            'Result(4) = JsonConvert.SerializeObject(ObjRspOAuth)
'            'Result(3) = JsonConvert.SerializeObject(ObjReqOAuth)
'            Result(2) = Access_Token
'            Result(1) = ""
'            Result(0) = "0"

'Skip:

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(wsAppName, wsAppVersion, Method, "", ex, Query)

'            Result(1) = ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Public Class cPPIdmReqPaidStatus
'        Public Produk As String = "PAYMENTPOINT"
'        Public Action As String = "ADVICE"
'        Public Detail As New cPPIdmReqPaidStatusDetail
'    End Class

'    Public Class cPPIdmReqPaidStatusDetail
'        Public PLU As String = ""
'        Public KodeBayar As String = ""
'        Public Key As String = "" 'token / oauth
'    End Class

'    Public Class cPPIdmRspPaidStatus
'        Public RespCode As String = "" '0 = sukses
'        Public RespDesc As String = ""
'        Public Detail As New cPPIdmRspPaidStatusDetail
'    End Class

'    Public Class cPPIdmRspPaidStatusDetail
'        Public Reference As String = "" 'nomor referensi transaksi sukses
'    End Class


'    Public Function PaymentPointMappingPLU(ByVal ProductName As String) As String

'        Dim Result As String = ""

'        Dim KnownProductName As String = ProductName.ToUpper

'        If KnownProductName.Contains("KLIK") And KnownProductName.Contains("INDOMARET") Then
'            KnownProductName = "KLIK"
'        End If
'        If KnownProductName.Contains("APKA") Then
'            KnownProductName = "APKA"
'        End If

'        Select Case KnownProductName
'            Case "KLIK"
'                Result = "20114741"
'            Case "APKA"
'                Result = "20114742"
'            Case "SPI"
'                Result = "SPI"
'            Case "KIGR"
'                Result = "KIGR"
'            Case Else

'        End Select

'        Return Result

'    End Function

'    <WebMethod()>
'    Public Function PaymentPointPaidStatus(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "PaymentPointPaidStatus"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim Keyword As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString
'            Dim ProductName As String = Param(2).ToString.Trim.ToUpper 'Klik atau APKA
'            Dim PaymentCode As String = Param(3).ToString.Trim 'Kode Bayar

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If ProductName = "" Or PaymentCode = "" Then
'                Result(1) = "Input ProductName dan PaymentCode"
'                GoTo Skip
'            End If

'            Dim ProductCode As String = PaymentPointMappingPLU(ProductName.ToUpper)
'            If ProductCode = "" Then
'                Result(1) = "ProductName " & ProductName & " tidak dikenali"
'                GoTo Skip
'            End If

'            Keyword = ProductName & " " & ProductCode & " " & PaymentCode

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "MULAI", Keyword)


'            Dim ObjSQL As New ClsSQL

'            Query = "Select count(Key1) From ByPassTrx"
'            Query &= " Where `Type` = 'CODPaymentCode' And Key1 = @ProductName And Key2 = @PaymentCode"
'            Query &= " And curdate() between ActiveDate and date_add(ActiveDate, interval Duration day)"

'            SqlParam = New Dictionary(Of String, String)
'            SqlParam.Add("@ProductName", ProductName)
'            SqlParam.Add("@PaymentCode", PaymentCode)

'            If ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam) > 0 Then
'                Process = True
'                Result(0) = "0"
'                Result(2) = "ByPass"
'                GoTo Skip
'            End If


'            Dim PaymentPointIndomaret As Boolean = False
'            If ProductName.ToUpper.Contains("KLIK") And ProductName.ToUpper.Contains("INDOMARET") Then
'                PaymentPointIndomaret = True
'            End If
'            If ProductName.ToUpper.Contains("APKA") Then
'                PaymentPointIndomaret = True
'            End If

'            If PaymentPointIndomaret Then

'                PaymentPointIndomaretInit()

'                Dim ObjReq As New cPPIdmReqPaidStatus

'                ObjReq.Detail.PLU = ProductCode
'                ObjReq.Detail.KodeBayar = PaymentCode

'                '=== token
'                Dim AlreadyResetToken As Boolean = False
'                Dim MustResetToken As Boolean = False

'UlangGetOAuthIndomaret:

'                Dim Access_Token As String = ""

'                Dim tParam() As Object
'                If MustResetToken = False Then
'                    ReDim tParam(1)
'                    tParam(0) = Param(0)
'                    tParam(1) = Param(1)
'                Else
'                    ReDim tParam(2)
'                    tParam(0) = Param(0)
'                    tParam(1) = Param(1)
'                    tParam(2) = "1"
'                    AlreadyResetToken = True
'                End If

'                Dim tResult As Object() = PaymentPointIndomaretGetOAuth(wsAppName, wsAppVersion, tParam)
'                If tResult(0).ToString = "0" Then
'                    Access_Token = tResult(2).ToString
'                Else
'                    Result(1) = tResult(1).ToString
'                    GoTo Skip
'                End If

'                ObjReq.Detail.Key = Access_Token
'                '=== token

'                Dim Parameter As String = JsonConvert.SerializeObject(ObjReq)

'                e.APIRequestLog(MCon, wsAppName, wsAppVersion, Method, wsUser, Parameter, mPPIdmURLPaidStatus, Keyword)

'                Dim pParam(2) As Object
'                pParam(0) = Parameter
'                pParam(1) = mPPIdmURLPaidStatus
'                pParam(2) = mPPIdmContentType

'                wsSlave.Url = ("" & ConfigurationManager.AppSettings("WServiceSlave.ServiceSlave")).Trim

'                Dim Response As String = wsSlave.PaymentPointIndomaretPaidStatusV2(pParam)

'                Dim ObjRsp As New cPPIdmRspPaidStatus

'                If Response <> "" Then
'                    e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsUser, Response, mPPIdmURLPaidStatus, Keyword)
'                    Try
'                        ObjRsp = JsonConvert.DeserializeObject(Of cPPIdmRspPaidStatus)(Response)

'                        Dim IsPaid As Boolean = (ObjRsp.RespCode = "0")

'                        If IsPaid = False Then
'                            Try
'                                '{"RespCode":"87","RespDesc":"APKA-Tagihan Sudah Dibayar","Detail":{"Reference":""}}
'                                IsPaid = (ObjRsp.RespCode = "87" And ObjRsp.RespDesc.ToLower.Contains("sudah") And ObjRsp.RespDesc.ToLower.Contains("bayar"))
'                            Catch ex As Exception
'                            End Try
'                        End If

'                        If IsPaid = False Then
'                            Try
'                                '{"RespCode":"6","RespDesc":"ECOM-Kode pesanan I96RQ21 sudah dibayar","Detail":{"Reference":""}}
'                                IsPaid = (ObjRsp.RespCode = "6" And ObjRsp.RespDesc.ToLower.Contains("sudah") And ObjRsp.RespDesc.ToLower.Contains("bayar"))
'                            Catch ex As Exception
'                            End Try
'                        End If

'                        If IsPaid = False Then
'                            Try
'                                '{"RespCode":"21","RespDesc":"PUSH MANUAL RBR","Detail":{"Reference":""}}
'                                IsPaid = (ObjRsp.RespCode = "21" And ObjRsp.RespDesc.ToLower.Contains("push") And ObjRsp.RespDesc.ToLower.Contains("rbr"))
'                            Catch ex As Exception
'                            End Try
'                        End If

'                        If IsPaid Then
'                            Result(0) = "0"
'                            Try
'                                Result(2) = ObjRsp.Detail.Reference
'                            Catch ex As Exception
'                                Result(2) = ""
'                            End Try
'                            If Result(2) = "" Then
'                                Result(2) = "PAID-" & PaymentCode
'                            End If
'                        Else
'                            Try
'                                If ObjRsp.RespCode.ToLower = "i21" _
'                                Or ObjRsp.RespDesc.ToLower.Contains("key ") Then
'                                    MustResetToken = True
'                                    GoTo UlangGetOAuthIndomaret
'                                End If
'                            Catch ex As Exception
'                            End Try
'                            Result(1) = ObjRsp.RespDesc & "." & ObjRsp.RespCode
'                        End If

'                    Catch ex As Exception
'                        e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, wsUser, ex, "", Keyword)
'                        Result(1) = ex.Message & vbCrLf & ex.StackTrace
'                    End Try
'                Else
'                    e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsUser, "No Response", mPPIdmURLPaidStatus, Keyword)
'                    Result(1) = "No Response"
'                End If 'dari If Response <> ""


'            ElseIf ProductName.ToUpper.Contains("SPI") Then
'                'PaymentPoint Indogrosir SPI - BEGIN

'                PaymentPointIndogrosirSPIInit()

'                Dim ObjReq As New cPPIgrSpiReqPaidStatus

'                If False Then
'                    Query = "Select TrackNum From TransactionDeliveryInfo"
'                    Query &= " Where CODPaymentBiller = 'SPI' and CODPaymentCode = @PaymentCode"

'                    SqlParam = New Dictionary(Of String, String)
'                    SqlParam.Add("@PaymentCode", PaymentCode)

'                    Dim TrackNumSPI As String = ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam)
'                    If TrackNumSPI = "" Or TrackNumSPI = "-1" Then
'                        Result(1) = "TrackNum tidak ditemukan (SPI " & PaymentCode & ")"
'                    End If

'                    ObjReq.awb = TrackNumSPI
'                End If

'                ObjReq.codpaymentcode = PaymentCode

'                ReDim CustomHeaders(1)
'                CustomHeaders(0) = "x-api-key|" & wsPassPPIgrSpi
'                CustomHeaders(1) = "Content-Type|" & "application/json"
'                Dim ParameterHeaders As String = ""
'                For h As Integer = 0 To CustomHeaders.Length - 1
'                    If ParameterHeaders <> "" Then
'                        ParameterHeaders &= ","
'                    End If
'                    ParameterHeaders &= CustomHeaders(h)
'                Next

'                Dim Parameter As String = JsonConvert.SerializeObject(ObjReq)

'                e.APIRequestLog(MCon, wsAppName, wsAppVersion, Method, wsUserPPIgrSpi, "Header:" & ParameterHeaders & "-" & "Body:" & Parameter, mPPIgrSpiURLPaidStatus, Keyword)

'                Dim ObjFungsi As New ClsFungsi

'                Dim Response As String = ObjFungsi.SendHTTP("", "", "", mPPIgrSpiURLPaidStatus, Parameter, wsUserAgent, Encoding.Default, 60, "", mPPIgrSpiContentType, True, CustomHeaders).Trim

'                Dim ObjRsp As New cPPIgrSpiRspPaidStatus

'                If Response <> "" Then
'                    e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsUserPPIgrSpi, Response, mPPIgrSpiURLPaidStatus, Keyword)

'                    Try
'                        ObjRsp = JsonConvert.DeserializeObject(Of cPPIgrSpiRspPaidStatus)(Response)

'                        Dim IsPaid As Boolean = (ObjRsp.data.status_id = "8")

'                        If IsPaid Then
'                            Result(0) = "0"
'                            Try
'                                Result(2) = ObjRsp.data.awb_code
'                            Catch ex As Exception
'                                Result(2) = ""
'                            End Try
'                            If Result(2) = "" Then
'                                Result(2) = "PAID-" & PaymentCode
'                            End If
'                        End If

'                    Catch ex As Exception
'                        e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, wsUserPPIgrSpi, ex, "", Keyword)
'                        Result(1) = ex.Message & vbCrLf & ex.StackTrace
'                    End Try

'                Else
'                    e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsUserPPIgrSpi, "No Response", mPPIgrSpiURLPaidStatus, Keyword)
'                    Result(1) = "No Response"

'                End If 'dari If Response <> ""

'                'PaymentPoint Indogrosir SPI - END


'            ElseIf ProductName.ToUpper.Contains("KIGR") Then
'                'PaymentPoint Indogrosir KIGR - BEGIN

'                PaymentPointIndogrosirKIGRInit()

'                Dim ObjReq As New cPPKIGRReqPaidStatus

'                If False Then
'                    Query = "Select TrackNum From TransactionDeliveryInfo"
'                    Query &= " Where CODPaymentBiller = 'KIGR' and CODPaymentCode = @PaymentCode"

'                    SqlParam = New Dictionary(Of String, String)
'                    SqlParam.Add("@PaymentCode", PaymentCode)

'                    Dim TrackNumKIGR As String = ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam)
'                    If TrackNumKIGR = "" Or TrackNumKIGR = "-1" Then
'                        Result(1) = "TrackNum tidak ditemukan (KIGR " & PaymentCode & ")"
'                    End If

'                    ObjReq.awb = TrackNumKIGR
'                End If

'                ObjReq.codpaymentcode = PaymentCode

'                ReDim CustomHeaders(1)
'                CustomHeaders(0) = "x-api-key|" & wsPassPPKIGR
'                CustomHeaders(1) = "Content-Type|" & "application/json"
'                Dim ParameterHeaders As String = ""
'                For h As Integer = 0 To CustomHeaders.Length - 1
'                    If ParameterHeaders <> "" Then
'                        ParameterHeaders &= ","
'                    End If
'                    ParameterHeaders &= CustomHeaders(h)
'                Next

'                Dim Parameter As String = JsonConvert.SerializeObject(ObjReq)

'                e.APIRequestLog(MCon, wsAppName, wsAppVersion, Method, wsUserPPKIGR, "Header:" & ParameterHeaders & "-" & "Body:" & Parameter, mPPKIGRURLPaidStatus, Keyword)

'                Dim ObjFungsi As New ClsFungsi

'                Dim Response As String = ObjFungsi.SendHTTP("", "", "", mPPKIGRURLPaidStatus, Parameter, wsUserAgent, Encoding.Default, 60, "", mPPKIGRContentType, True, CustomHeaders).Trim

'                Dim ObjRsp As New cPPKIGRRspPaidStatus

'                If Response <> "" Then
'                    e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsUserPPKIGR, Response, mPPKIGRURLPaidStatus, Keyword)

'                    Try
'                        ObjRsp = JsonConvert.DeserializeObject(Of cPPKIGRRspPaidStatus)(Response)

'                        Dim IsPaid As Boolean = (ObjRsp.data.status_id = "8")

'                        If IsPaid Then
'                            Result(0) = "0"
'                            Try
'                                Result(2) = ObjRsp.data.awb_code
'                            Catch ex As Exception
'                                Result(2) = ""
'                            End Try
'                            If Result(2) = "" Then
'                                Result(2) = "PAID-" & PaymentCode
'                            End If
'                        End If

'                    Catch ex As Exception
'                        e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, wsUserPPKIGR, ex, "", Keyword)
'                        Result(1) = ex.Message & vbCrLf & ex.StackTrace
'                    End Try

'                Else
'                    e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsUserPPKIGR, "No Response", mPPKIGRURLPaidStatus, Keyword)
'                    Result(1) = "No Response"

'                End If 'dari If Response <> ""

'                'PaymentPoint Indogrosir KIGR - END


'            Else

'                Result(1) = "Tidak tersedia " & ProductName

'            End If 'dari If PaymentPointIndomaret


'            If Result(0).ToString = "0" Then
'                Process = True
'            End If

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "SELESAI " & Process.ToString, Keyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Public mPPIgrSpiURLPaidStatus As String = ""
'    Public mPPIgrSpiContentType As String = ""
'    Public wsUserPPIgrSpi As String = ""
'    Public wsPassPPIgrSpi As String = ""

'    Public Sub PaymentPointIndogrosirSPIInit()

'        If mPPIgrSpiURLPaidStatus = "" Then
'            Try
'                mPPIgrSpiURLPaidStatus = ("" & ConfigurationManager.AppSettings("PaymentPointIgrSpiUrlPaidStatus")).Trim
'            Catch ex As Exception
'                mPPIgrSpiURLPaidStatus = ""
'            End Try
'            If mPPIgrSpiURLPaidStatus = "" Then
'                mPPIgrSpiURLPaidStatus = "x"
'            End If
'        End If

'        If mPPIgrSpiContentType = "" Then
'            mPPIgrSpiContentType = "application/json"
'        End If

'        If wsUserPPIgrSpi = "" Then
'            Try
'                wsUserPPIgrSpi = ("" & ConfigurationManager.AppSettings("PaymentPointIgrSpiWsUser")).Trim
'            Catch ex As Exception
'                wsUserPPIgrSpi = ""
'            End Try
'            If wsUserPPIgrSpi = "" Then
'                wsUserPPIgrSpi = "x"
'            End If
'        End If

'        If wsPassPPIgrSpi = "" Then
'            Try
'                wsPassPPIgrSpi = ("" & ConfigurationManager.AppSettings("PaymentPointIgrSpiWsPassword")).Trim
'            Catch ex As Exception
'                wsPassPPIgrSpi = ""
'            End Try
'            If wsPassPPIgrSpi = "" Then
'                wsPassPPIgrSpi = "x"
'            End If
'        End If

'    End Sub

'    Public Class cPPIgrSpiReqPaidStatus
'        Public awb As String = ""
'        Public codpaymentcode As String = ""
'    End Class

'    Public Class cPPIgrSpiRspPaidStatus
'        Public response_code As String = "" '200 = ada response
'        Public response_message As String = "" 'pesan berhasil / error
'        Public data As New cPPIgrSpiRspPaidStatusDetail
'    End Class

'    Public Class cPPIgrSpiRspPaidStatusDetail
'        Public awb_code As String = ""
'        Public status_id As String = "" '8 = sudah paid
'        Public description As String = ""
'    End Class


'    Public mPPKIGRURLPaidStatus As String = ""
'    Public mPPKIGRContentType As String = ""
'    Public wsUserPPKIGR As String = ""
'    Public wsPassPPKIGR As String = ""

'    Public Sub PaymentPointIndogrosirKIGRInit()

'        If mPPKIGRURLPaidStatus = "" Then
'            Try
'                mPPKIGRURLPaidStatus = ("" & ConfigurationManager.AppSettings("PaymentPointKIGRUrlPaidStatus")).Trim
'            Catch ex As Exception
'                mPPKIGRURLPaidStatus = ""
'            End Try
'            If mPPKIGRURLPaidStatus = "" Then
'                mPPKIGRURLPaidStatus = "x"
'            End If
'        End If

'        If mPPKIGRContentType = "" Then
'            mPPKIGRContentType = "application/json"
'        End If

'        If wsUserPPKIGR = "" Then
'            Try
'                wsUserPPKIGR = ("" & ConfigurationManager.AppSettings("PaymentPointKIGRWsUser")).Trim
'            Catch ex As Exception
'                wsUserPPKIGR = ""
'            End Try
'            If wsUserPPKIGR = "" Then
'                wsUserPPKIGR = "x"
'            End If
'        End If

'        If wsPassPPKIGR = "" Then
'            Try
'                wsPassPPKIGR = ("" & ConfigurationManager.AppSettings("PaymentPointKIGRWsPassword")).Trim
'            Catch ex As Exception
'                wsPassPPKIGR = ""
'            End Try
'            If wsPassPPKIGR = "" Then
'                wsPassPPKIGR = "x"
'            End If
'        End If

'    End Sub

'    Public Class cPPKIGRReqPaidStatus
'        Public awb As String = ""
'        Public codpaymentcode As String = ""
'    End Class

'    Public Class cPPKIGRRspPaidStatus
'        Public response_code As String = "" '200 = ada response
'        Public response_message As String = "" 'pesan berhasil / error
'        Public data As New cPPKIGRRspPaidStatusDetail
'    End Class

'    Public Class cPPKIGRRspPaidStatusDetail
'        Public awb_code As String = ""
'        Public status_id As String = "" '8 = sudah paid
'        Public description As String = ""
'    End Class


'    <WebMethod()>
'    Public Function CheckPINStatus(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "CheckPINStatus"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString
'            Dim DeliveryId As String = Param(2).ToString.Trim.ToUpper 'DeliveryId milik ThirdParty
'            Dim PINType As String = Param(3).ToString.Trim.ToUpper 'PICKUP, CANCEL, RETURN, KEEP
'            Dim PINString As String = Param(4).ToString.Trim.ToUpper 'PIN yang dimaksud

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If DeliveryId = "" Or PINType = "" Then
'                Result(1) = "Input DeliveryId dan PINType"
'                GoTo Skip
'            End If

'            Dim Keyword As String = DeliveryId & " " & PINType

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "MULAI", Keyword)

'            Dim ObjSQL As New ClsSQL

'            Dim ThirdPartyList As String = CustomAccount3rdPartyDeliveryMan
'            If CustomAccount3rdPartyIndopaketMotor <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketMotor
'            End If
'            If CustomAccount3rdPartyIndopaketMobil <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketMobil
'            End If
'            If CustomAccount3rdPartyIndopaketInstantMotor <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketInstantMotor
'            End If
'            If CustomAccount3rdPartyIndopaketInstantMobil <> "" Then
'                ThirdPartyList &= "," & CustomAccount3rdPartyIndopaketInstantMobil
'            End If

'            Query = "Select o.TrackNum, t.AddInfo as tAddInfo"
'            Query &= " From AutoOrderThirdPartyTransaction o"
'            Query &= " Inner Join `Transaction` t on (t.TrackNum = o.TrackNum)"
'            Query &= " Where o.ThirdParty in (" & ThirdPartyList & ")"
'            Query &= " And o.dOrderNo = @DeliveryId"

'            SqlParam = New Dictionary(Of String, String)
'            SqlParam.Add("@DeliveryId", DeliveryId)

'            Dim dtTransaction As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)

'            If dtTransaction Is Nothing Then
'                Result(1) = "Gagal query TrackNum (" & DeliveryId & ")"
'                GoTo Skip
'            End If
'            If dtTransaction.Rows.Count < 1 Then
'                Result(1) = "TrackNum tidak ditemukan (" & DeliveryId & ")"
'                GoTo Skip
'            End If

'            Dim TrackNum As String = dtTransaction.Rows(0).Item("TrackNum").ToString
'            Dim TAddInfo As String = dtTransaction.Rows(0).Item("tAddInfo").ToString.ToUpper

'            Query = "Select TrackNum, OrderNo, OrderTypeDetail"
'            Query &= " , PINPickup, PINPickupDraft, PINCancel, PINCancelDraft"
'            Query &= " , PINReturn, PINReturnDraft, PINKeep, PINKeepDraft"
'            Query &= " , case when AddInfo like '%PICKUPCANCEL=Y%' then 1 else 0 end as IsPickupCancel"
'            Query &= " From TransactionDeliveryInfo Where TrackNum = @TrackNum"

'            SqlParam = New Dictionary(Of String, String)
'            SqlParam.Add("@TrackNum", TrackNum)

'            Dim dtDeliveryInfo As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)

'            If dtDeliveryInfo Is Nothing Then
'                Result(1) = "Gagal query DeliveryInfo (" & DeliveryId & ")"
'                GoTo Skip
'            End If

'            If dtDeliveryInfo.Rows.Count < 1 Then
'                Result(1) = "DeliveryInfo tidak ditemukan (" & DeliveryId & ")"
'                GoTo Skip
'            End If

'            If TAddInfo.Contains("BYPSPINDLV=Y") Then
'                'bypass pengecekan PIN
'                Process = True

'                Result(0) = "0"
'                Result(1) = ""

'                GoTo Skip
'            End If


'            Dim OrderTypeDetail As String = ""
'            Try
'                OrderTypeDetail = dtDeliveryInfo.Rows(0).Item("OrderTypeDetail").ToString.ToUpper
'                If OrderTypeDetail.Contains("INDO") And OrderTypeDetail.Contains("GROSIR") Then
'                    OrderTypeDetail = "INDOGROSIR"
'                ElseIf OrderTypeDetail.Contains("OTHER") Then
'                    OrderTypeDetail = ""
'                Else
'                    OrderTypeDetail = "INDOMARET"
'                End If
'            Catch ex As Exception
'                OrderTypeDetail = ""
'            End Try

'            Dim DeliveryInfoOrderNo As String = ""
'            Try
'                DeliveryInfoOrderNo = dtDeliveryInfo.Rows(0).Item("OrderNo").ToString.ToUpper
'            Catch ex As Exception
'                DeliveryInfoOrderNo = ""
'            End Try


'            Dim IsPickupCancel As Boolean = False
'            Dim PINTypeDescription As String = ""

'            Select Case PINType
'                Case "PICKUP"

'                    Dim dtPickupCancel As DataTable = dtDeliveryInfo.DefaultView.ToTable
'                    'dtPickupCancel.DefaultView.RowFilter = "PINPickup = PINPickupDraft And IsPickupCancel= 1"
'                    dtPickupCancel.DefaultView.RowFilter = "IsPickupCancel= 1"
'                    If dtPickupCancel.DefaultView.Count > 0 Then
'                        IsPickupCancel = True
'                    End If

'                    'dtDeliveryInfo.DefaultView.RowFilter = "PINPickup <> PINPickupDraft"
'                    dtDeliveryInfo.DefaultView.RowFilter = "PINPickup <> PINPickupDraft And IsPickupCancel <> 1"

'                    PINTypeDescription = "Ambil"

'                Case "CANCEL"
'                    dtDeliveryInfo.DefaultView.RowFilter = "PINCancel <> PINCancelDraft"

'                    PINTypeDescription = "Batal"

'                Case "RETURN"
'                    dtDeliveryInfo.DefaultView.RowFilter = "PINReturn <> PINReturnDraft"

'                    PINTypeDescription = "Retur Bulky"

'                Case "KEEP"
'                    dtDeliveryInfo.DefaultView.RowFilter = "PINKeep <> PINKeepDraft"

'                    PINTypeDescription = "Titip"

'                Case Else
'                    Result(1) = "Belum tersedia (" & PINType & ")"
'                    GoTo Skip

'            End Select

'            If dtDeliveryInfo.DefaultView.Count < 1 Then

'                If PINType = "PICKUP" And IsPickupCancel Then
'                    Result(1) = DeliveryId & " sudah di-cancel"
'                    Select Case OrderTypeDetail
'                        Case "INDOMARET"
'                            Result(1) &= " oleh Toko Indomaret"
'                        Case "INDOGROSIR"
'                            Result(1) &= " oleh Indogrosir"
'                        Case Else
'                            'tanpa keterangan lokasi
'                    End Select
'                    If DeliveryInfoOrderNo <> "" Then
'                        Result(1) &= " (" & DeliveryInfoOrderNo & ")"
'                    End If
'                Else
'                    'response default
'                    Result(1) = DeliveryId & " belum proses PIN " & PINTypeDescription

'                    Select Case OrderTypeDetail
'                        Case "INDOMARET"
'                            Result(1) &= " di Toko Indomaret"
'                        Case "INDOGROSIR"
'                            Result(1) &= " di Indogrosir"
'                        Case Else
'                            'tanpa keterangan lokasi
'                    End Select
'                End If

'                GoTo Skip

'            End If 'dari If dtDeliveryInfo.DefaultView.Count < 1

'            Process = True

'            Result(0) = "0"
'            Result(1) = ""

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "SELESAI " & Process.ToString, Keyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Public Class cSyncCluster
'        Public data As cSyncClusterData
'        Public timestamp As String = ""
'        Public signature As String = ""
'    End Class

'    Public Class cSyncClusterData
'        Public id As String = ""
'        Public clusterCode As String = ""
'        Public clusterName As String = ""
'        Public postalCodeDescription As String = ""
'        Public store As cSyncClusterStore()
'        Public deletedStore As cSyncClusterStore()
'        Public driver As cSyncClusterDriver()
'        Public deletedDriver As cSyncClusterDriver()
'        Public active As Boolean = False
'        Public coordinatorNik As String = ""
'        Public coordinatorName As String = ""
'    End Class

'    Public Class cSyncClusterStore
'        Public storeCode As String = ""
'        Public storeName As String = ""
'    End Class

'    Public Class cSyncClusterDriver
'        Public nik As String = ""
'        Public name As String = ""
'    End Class

'    <WebMethod()>
'    Public Function SyncCluster(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "SyncCluster"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString

'            Dim Request As String = Param(2).ToString
'            Request = Request.Trim

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If Request = "" Then
'                Result(1) = "Input request"
'                GoTo Skip
'            End If

'            Dim Keyword As String = Format(Date.Now, "yyMMddHHmmss")

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "MULAI " & Request, Keyword)


'            Dim ObjRequest As cSyncCluster = JsonConvert.DeserializeObject(Of cSyncCluster)(Request)

'            If ObjRequest Is Nothing Then
'                Result(1) = "Gagal convert parameter request"
'                GoTo Skip
'            End If


'            Dim dtQuery As New DataTable
'            dtQuery.Columns.Add("SqlQuery")
'            dtQuery.Columns.Add("Result")

'            Dim ObjSQL As New ClsSQL

'            Dim ClusterId As String = ObjRequest.data.id
'            Dim ClusterCode As String = ObjRequest.data.clusterCode
'            Dim ClusterName As String = ObjRequest.data.clusterName
'            Dim PostalCodeDesc As String = WService.StrSQLRemoveChar(ObjRequest.data.postalCodeDescription)
'            Dim IsActive As Integer = IIf(ObjRequest.data.active, 1, 0)
'            Dim CoordinatorId As String = ObjRequest.data.coordinatorNik
'            Dim CoordinatorName As String = WService.StrSQLRemoveChar(ObjRequest.data.coordinatorName)

'            Dim sb As New StringBuilder
'            sb.Append("Insert Into DmsCluster (")
'            sb.Append(" ClusterId, Code, Name, PostalCodeDesc")
'            sb.Append(" , CoordinatorId, CoordinatorName, IsActive")
'            sb.Append(" , AddTime, AddUser")
'            sb.Append(" ) values (")
'            sb.Append(" '" & ClusterId & "', '" & ClusterCode & "', '" & ClusterName & "', '" & PostalCodeDesc & "'")
'            sb.Append(" , '" & CoordinatorId & "', '" & CoordinatorName & "', '" & IsActive & "'")
'            sb.Append(" , now(), 'SYNC_ADD'")
'            sb.Append(" ) on duplicate key update")
'            sb.Append(" UpdTime = now(), UpdUser = 'SYNC_UPD'")
'            sb.Append(" , ClusterId = '" & ClusterId & "', Code = '" & ClusterCode & "', Name = '" & ClusterName & "', PostalCodeDesc = '" & PostalCodeDesc & "'")
'            sb.Append(" , CoordinatorId = '" & CoordinatorId & "', CoordinatorName = '" & CoordinatorName & "', IsActive = '" & IsActive & "'")

'            dtQuery.Rows.Add()
'            dtQuery.Rows(dtQuery.Rows.Count - 1).Item("SqlQuery") = sb.ToString
'            dtQuery.AcceptChanges()

'            Try
'                For i As Integer = 0 To ObjRequest.data.store.Length - 1
'                    sb = New StringBuilder
'                    sb.Append("Insert Ignore Into DmsClusterStore (")
'                    sb.Append(" ClusterId, Code, Name, UpdTime, UpdUser")
'                    sb.Append(" ) values (")
'                    sb.Append(" '" & ClusterId & "', '" & ObjRequest.data.store(i).storeCode & "', '" & WService.StrSQLRemoveChar(ObjRequest.data.store(i).storeName) & "', now(), 'SYNC_STO'")
'                    sb.Append(" )")

'                    dtQuery.Rows.Add()
'                    dtQuery.Rows(dtQuery.Rows.Count - 1).Item("SqlQuery") = sb.ToString
'                    dtQuery.AcceptChanges()
'                Next
'            Catch ex As Exception
'            End Try

'            Try
'                For i As Integer = 0 To ObjRequest.data.deletedStore.Length - 1
'                    sb = New StringBuilder
'                    sb.Append("Delete From DmsClusterStore Where ClusterId = '" & ClusterId & "' And Code = '" & ObjRequest.data.deletedStore(i).storeCode & "'")

'                    dtQuery.Rows.Add()
'                    dtQuery.Rows(dtQuery.Rows.Count - 1).Item("SqlQuery") = sb.ToString
'                    dtQuery.AcceptChanges()
'                Next
'            Catch ex As Exception
'            End Try


'            Try
'                For i As Integer = 0 To ObjRequest.data.driver.Length - 1
'                    sb = New StringBuilder
'                    sb.Append("Insert Ignore Into DmsClusterDriver (")
'                    sb.Append(" ClusterId,  DriverId, Name, UpdTime, UpdUser")
'                    sb.Append(" ) values (")
'                    sb.Append(" '" & ClusterId & "', '" & ObjRequest.data.driver(i).nik & "', '" & WService.StrSQLRemoveChar(ObjRequest.data.driver(i).name) & "', now(), 'SYNC_DRV'")
'                    sb.Append(" )")

'                    dtQuery.Rows.Add()
'                    dtQuery.Rows(dtQuery.Rows.Count - 1).Item("SqlQuery") = sb.ToString
'                    dtQuery.AcceptChanges()
'                Next
'            Catch ex As Exception
'            End Try

'            Try
'                For i As Integer = 0 To ObjRequest.data.deletedDriver.Length - 1
'                    sb = New StringBuilder
'                    sb.Append("Delete From DmsClusterDriver Where ClusterId = '" & ClusterId & "' And DriverId = '" & ObjRequest.data.deletedDriver(i).nik & "'")

'                    dtQuery.Rows.Add()
'                    dtQuery.Rows(dtQuery.Rows.Count - 1).Item("SqlQuery") = sb.ToString
'                    dtQuery.AcceptChanges()
'                Next
'            Catch ex As Exception
'            End Try


'            For i As Integer = 0 To dtQuery.Rows.Count - 1
'                Query = dtQuery.Rows(i).Item("SqlQuery").ToString
'                dtQuery.Rows(i).Item("Result") = IIf(ObjSQL.ExecNonQuery(MCon, Query), "1", "0")
'                dtQuery.AcceptChanges()
'            Next


'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = ""

'            Process = True

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "SELESAI " & Process.ToString, Keyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Dim RePushStatusManualUser As String = ""
'    Dim RePushStatusManualPassword As String = ""
'    Dim RePushStatusManualUserId As String = ""
'    Dim RePushStatusManualDeliveryIdList As String = ""

'    <WebMethod()>
'    Public Function RePushStatusManual(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object() = CreateResult()

'        Dim Method As String = "RePushStatusManual"

'        Dim MCon As MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString
'            Dim UserId As String = Param(2).ToString
'            Dim TrackNumList As String = Param(3).ToString 'TrackNum1|TrackNum2|dst

'            Dim LogKeyword As String = UserId & Format(Date.Now, "yyMMddHHmmss")

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            WService.CreateRequestLog(wsAppName, wsAppVersion, Method, User, UserId, Param, LogKeyword)

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            Dim ObjSQL As New ClsSQL

'            'sp untuk dapatin daftar DeliveryId
'            Query = "call sp_dms_repushstatus ('" & TrackNumList & "')"
'            RePushStatusManualDeliveryIdList = ObjSQL.ExecScalar(MCon, Query)
'            'sudah dalam format "{"orderId":"DMS-111"},{"orderId":"DMS-222"},dst"


'            RePushStatusManualUser = User
'            RePushStatusManualPassword = Password

'            Dim t As New Threading.Thread(AddressOf RePushStatusManualCore)
'            t.Start()


'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = RePushStatusManualDeliveryIdList.Replace("""", "").Replace("orderId", "").Replace(":", "").Replace("{", "").Replace("}", "")

'Skip:

'            WService.CreateResponseLog(wsAppName, wsAppVersion, Method, User, Result, , LogKeyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(wsAppName, wsAppVersion, Method, User, ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function

'    Private Sub RePushStatusManualCore()

'        Dim Result As Object() = CreateResult()

'        Dim Method As String = "RePushStatusManualCore"

'        Dim MCon As MySqlConnection

'        Dim Query As String = ""

'        Dim e As New ClsError

'        Try
'            MCon = MasterMCon.Clone
'            MCon.Open()

'            'config api DMS
'            Dim mRePushStatusUrl As String = ""
'            Try
'                mRePushStatusUrl = ("" & ConfigurationManager.AppSettings("DeliveryManAPIRePushStatusUrl")).Trim
'            Catch ex As Exception
'                mRePushStatusUrl = ""
'            End Try
'            If mRePushStatusUrl = "" Then
'                mRePushStatusUrl = "x"
'            End If

'            Dim mRePushStatusLimit As String = ""
'            Try
'                mRePushStatusLimit = ("" & ConfigurationManager.AppSettings("DeliveryManAPIRePushStatusLimit")).Trim
'            Catch ex As Exception
'                mRePushStatusLimit = ""
'            End Try
'            If mRePushStatusLimit = "" Then
'                mRePushStatusLimit = "20"
'            End If
'            Dim nRePushLimit As Integer = CInt(mRePushStatusLimit)


'            Dim Access_Token As String = ""

'            Dim tParam(1) As Object
'            tParam(0) = RePushStatusManualUser
'            tParam(1) = RePushStatusManualPassword

'            Dim tResult As Object() = GetOAuth(wsAppName, wsAppVersion, tParam)
'            If tResult(0).ToString = "0" Then
'                Access_Token = tResult(2).ToString
'            Else
'                Result(1) = tResult(1).ToString
'                GoTo Skip
'            End If

'            ReDim CustomHeaders(0)
'            CustomHeaders(0) = "X-API-Key|" & Access_Token

'            Dim ParameterHeaders As String = ""
'            For h As Integer = 0 To CustomHeaders.Length - 1
'                If ParameterHeaders <> "" Then
'                    ParameterHeaders &= ","
'                End If
'                ParameterHeaders &= CustomHeaders(h)
'            Next


'            Dim RePushStatusManualDeliveryIdListSplit As String() = RePushStatusManualDeliveryIdList.Split(",")


'            Dim j As Integer = 1
'            Dim RePushList As String = ""

'            Dim ObjFungsi As New ClsFungsi

'            For i As Integer = 0 To RePushStatusManualDeliveryIdListSplit.Length - 1

'                If RePushStatusManualDeliveryIdListSplit(i) <> "" Then

'                    If RePushList <> "" Then
'                        RePushList &= ","
'                    End If
'                    RePushList &= RePushStatusManualDeliveryIdListSplit(i)

'                    j += 1

'                    If (j > nRePushLimit) Or (i = RePushStatusManualDeliveryIdListSplit.Length - 1) Then
'                        'push data ke api, ketika:
'                        'mencapai jumlah batas data per 1 x push
'                        'atau sudah di akhir data

'                        Try
'                            Dim LogKeyword As String = RePushStatusManualUserId & Format(Date.Now, "yyMMddHHmmss")

'                            RePushList = "[" & RePushList & "]"

'                            e.APIRequestLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "Header:" & ParameterHeaders & "-" & "Body:" & RePushList, mRePushStatusUrl, LogKeyword)

'                            Dim Response As String = ""
'                            Response = ObjFungsi.SendHTTP("", "", "", mRePushStatusUrl, RePushList, "", Encoding.Default, "60", "", mContentType, True, CustomHeaders)
'                            Response = ("" & Response).Trim

'                            If Response <> "" Then
'                                e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, Response, mRePushStatusUrl, LogKeyword)
'                            Else
'                                e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "No Response", mRePushStatusUrl, LogKeyword)
'                            End If

'                        Catch ex As Exception
'                            e.ErrorLog(wsAppName, wsAppVersion, Method, User, ex, RePushList)

'                        End Try

'                        'reset data dan jumlah data yang akan di-push
'                        j = 1
'                        RePushList = ""

'                        System.Threading.Thread.Sleep(500)
'                    End If

'                End If 'dari If RePushStatusManualDeliveryIdListSplit(i) <> ""

'            Next

'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = ""

'Skip:

'        Catch ex As Exception
'            e.ErrorLog(wsAppName, wsAppVersion, Method, User, ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'    End Sub

'    <WebMethod()>
'    Public Function RePushStatusManualV2(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object() = CreateResult()

'        Dim Method As String = "RePushStatusManualV2"

'        Dim MCon As MySqlConnection

'        Dim Query As String = ""

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString
'            Dim UserId As String = Param(2).ToString
'            Dim TrackNumList As String = Param(3).ToString 'TrackNum1|TrackNum2|dst

'            Dim LogKeyword As String = UserId & Format(Date.Now, "yyMMddHHmmss")

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            WService.CreateRequestLog(wsAppName, wsAppVersion, Method, User, UserId, Param, LogKeyword)

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            Dim ObjSQL As New ClsSQL

'            'sp untuk dapatin daftar DeliveryId
'            Query = "call sp_dms_repushstatus ('" & TrackNumList & "')"
'            RePushStatusManualDeliveryIdList = ObjSQL.ExecScalar(MCon, Query)
'            'sudah dalam format "{"orderId":"DMS-111"},{"orderId":"DMS-222"},dst"

'            Query = "call sp_dms_repushstatusV2 ('" & TrackNumList & "')"
'            Dim dtQuery As DataTable = ObjSQL.ExecDatatable(MCon, Query)
'            'dalam format datatable untuk balikan ke webreport

'            RePushStatusManualUser = User
'            RePushStatusManualPassword = Password

'            Dim t As New Threading.Thread(AddressOf RePushStatusManualCore)
'            t.Start()

'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = WService.ConvertDatatableToString(dtQuery)

'Skip:

'            WService.CreateResponseLog(wsAppName, wsAppVersion, Method, User, Result, , LogKeyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(wsAppName, wsAppVersion, Method, User, ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Public Class cPushTrackingPUCResp
'        Public status As String = "" '00 = success
'    End Class

'    <WebMethod()>
'    Public Function sendTrackPUC(ByVal AppName As String, ByVal AppVersion As String) As Boolean

'        Return PushTrackingPUC(AppName, AppVersion)

'    End Function

'    Private Function PushTrackingPUC(Optional ByVal AppName As String = "", Optional ByVal AppVersion As String = "") As Boolean

'        Dim Method As String = "PushTrackingPUC"

'        Dim Result As Boolean = False

'        Dim MCon As New MySqlConnection
'        Dim Query As String = ""

'        Try
'            MCon = MasterMCon.Clone

'            Dim e As New ClsError
'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, "START")


'            Dim nLimit As Integer = 0
'            Try
'                nLimit = CInt("" & ConfigurationManager.AppSettings("DeliveryManAPICancelByStoreLimit"))
'                If nLimit.ToString.Trim = "" Then
'                    nLimit = 0
'                End If
'                If IsNumeric(nLimit) = False Then
'                    nLimit = 0
'                End If
'            Catch ex As Exception
'                nLimit = 0
'            End Try
'            If nLimit < 1 Then
'                nLimit = 60
'            End If


'            Query = "Select TrackId, TrackNum From PushTracking"
'            Query &= " Where Account = 'DMS' And TrackStatus = 'PUC' And PushStatus = '0'"
'            Query &= " Order By RetryPush, TrackTime"
'            Query &= " Limit " & nLimit

'            Dim ObjSQL As New ClsSQL
'            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)
'            If Not dtQuery Is Nothing Then
'                If dtQuery.Rows.Count > 0 Then

'                    Dim ObjFungsi As New ClsFungsi

'                    Dim Access_Token As String = ""

'                    Dim tParam(1) As Object
'                    tParam(0) = "system"
'                    tParam(1) = "#sY5t3M"
'                    Dim tResult As Object() = GetOAuth(wsAppName, wsAppVersion, tParam)
'                    If tResult(0).ToString = "0" Then
'                        Access_Token = tResult(2).ToString
'                    Else
'                        Access_Token = "x"
'                    End If

'                    ReDim CustomHeaders(0)
'                    CustomHeaders(0) = "X-API-Key|" & Access_Token

'                    Dim ParameterHeaders As String = ""
'                    For h As Integer = 0 To CustomHeaders.Length - 1
'                        If ParameterHeaders <> "" Then
'                            ParameterHeaders &= ","
'                        End If
'                        ParameterHeaders &= CustomHeaders(h)
'                    Next

'                    Dim mURLPUC As String = ""
'                    Try
'                        mURLPUC = ("" & ConfigurationManager.AppSettings("DeliveryManAPICancelByStoreUrl")).Trim
'                    Catch ex As Exception
'                        mURLPUC = ""
'                    End Try
'                    If mURLOrder = "" Then
'                        mURLPUC = "x"
'                    End If

'                    For i As Integer = 0 To dtQuery.Rows.Count - 1

'                        Dim TrackId As String = dtQuery.Rows(i).Item("TrackId").ToString
'                        Dim TrackNum As String = dtQuery.Rows(i).Item("TrackNum").ToString
'                        Dim WithError As String = "0" 'untuk penanda bila ada kegagalan

'                        Query = "Select ThirdParty, dOrderNo"
'                        Query &= " From AutoOrderThirdPartyTransaction"
'                        Query &= " Where TrackNum = @TrackNum"
'                        Query &= " And Status2 = 1"

'                        SqlParam = New Dictionary(Of String, String)
'                        SqlParam.Add("@TrackNum", TrackNum)

'                        Dim dtTemp As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)

'                        If Not (dtTemp Is Nothing) Then
'                            If dtTemp.Rows.Count > 0 Then

'                                For j As Integer = 0 To dtTemp.Rows.Count - 1
'                                    If CustomAccount3rdPartyDeliveryMan.Contains(dtTemp.Rows(j).Item("ThirdParty").ToString) Then
'                                        Dim dOrderNo As String = dtTemp.Rows(j).Item("dOrderNo").ToString

'                                        Dim Parameter As String = "[{""orderId"":""" & dOrderNo & """}]"

'                                        e.APIRequestLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, "Header:" & ParameterHeaders & "-Body:" & Parameter, mURLPUC, TrackNum & " " & dOrderNo)

'                                        Dim Response As String = ObjFungsi.SendHTTP("", "", "", mURLPUC, Parameter, "", Encoding.Default, "60", "", mContentType, True, CustomHeaders)
'                                        Response = ("" & Response).Trim

'                                        If Response <> "" Then

'                                            e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, Response, mURLPUC, TrackNum & " " & dOrderNo)

'                                            Dim ObjResponse As cPushTrackingPUCResp
'                                            Try
'                                                ObjResponse = JsonConvert.DeserializeObject(Of cPushTrackingPUCResp)(Response)

'                                                If ObjResponse.status = "00" Then
'                                                    'berhasil
'                                                Else
'                                                    WithError = "8"
'                                                End If
'                                            Catch ex As Exception
'                                                WithError = "9"
'                                                e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, ex, TrackNum & " " & dOrderNo)
'                                            End Try

'                                        Else
'                                            WithError = "9"
'                                            e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, "No Response", mURLPUC, TrackNum & " " & dOrderNo)

'                                        End If 'dari If Response <> ""

'                                    Else
'                                        WithError = "7"

'                                    End If 'dari If CustomAccount3rdPartyDeliveryMan.Contains(dtTemp.Rows(j).Item("ThirdParty").ToString)
'                                Next

'                            End If
'                        End If 'dari If Not (dtTemp Is Nothing)


'                        'apapun hasil-nya, anggap berhasil
'                        Query = "Update PushTracking"
'                        Query &= " Set `PushStatus` = '1'"
'                        If WithError <> "0" Then
'                            Query &= " , `RetryPush` = @WithError"
'                        End If
'                        Query &= " , UpdTime = now(), UpdUser = @AppName"
'                        Query &= " Where TrackId = @TrackId"
'                        Query &= " And PushStatus = '0'"

'                        SqlParam = New Dictionary(Of String, String)
'                        SqlParam.Add("@AppName", Strings.Left(wsAppName, 20))
'                        SqlParam.Add("@TrackId", TrackId)
'                        If WithError <> "0" Then
'                            SqlParam.Add("@WithError", WithError)
'                        End If

'                        ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

'                    Next ' dari For i As Integer = 0 To dtQuery.Rows.Count - 1

'                End If
'            End If 'dari If Not dtQuery Is Nothing

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, "FINISH")

'            Result = True

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(wsAppName, wsAppVersion, Method, wsAppName, ex, Query)

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Public Class cPushTrackingRCTReq
'        Public orderId As String = ""
'        Public isStoreReceipt As Boolean = True
'        Public receiptTimestamp As String = ""
'    End Class

'    Public Class cPushTrackingRCTResp
'        Public status As String = "" '00 = success
'    End Class

'    <WebMethod()>
'    Public Function sendTrackRCT(ByVal AppName As String, ByVal AppVersion As String) As Boolean

'        Return PushTrackingRCT(AppName, AppVersion)

'    End Function

'    Private Function PushTrackingRCT(Optional ByVal AppName As String = "", Optional ByVal AppVersion As String = "") As Boolean

'        Dim Method As String = "PushTrackingRCT"

'        Dim Result As Boolean = False

'        Dim MCon As New MySqlConnection
'        Dim Query As String = ""

'        Try
'            MCon = MasterMCon.Clone

'            Dim e As New ClsError
'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, "START")


'            Dim nLimit As Integer = 0
'            Try
'                nLimit = CInt("" & ConfigurationManager.AppSettings("DeliveryManAPIUpdateStoreReceiptLimit"))
'                If nLimit.ToString.Trim = "" Then
'                    nLimit = 0
'                End If
'                If IsNumeric(nLimit) = False Then
'                    nLimit = 0
'                End If
'            Catch ex As Exception
'                nLimit = 0
'            End Try
'            If nLimit < 1 Then
'                nLimit = 60
'            End If


'            Query = "Select TrackId, TrackNum, OrderNo, cast(TrackTime as char) as TrackTime"
'            Query &= " From PushTracking"
'            Query &= " Where Account = 'DMS' And TrackStatus = 'RCT' And PushStatus = '0'"
'            Query &= " Order By RetryPush, TrackTime"
'            Query &= " Limit " & nLimit

'            Dim ObjSQL As New ClsSQL
'            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)
'            If Not dtQuery Is Nothing Then
'                If dtQuery.Rows.Count > 0 Then

'                    Dim ObjFungsi As New ClsFungsi

'                    Dim Access_Token As String = ""

'                    Dim tParam(1) As Object
'                    tParam(0) = "system"
'                    tParam(1) = "#sY5t3M"
'                    Dim tResult As Object() = GetOAuth(wsAppName, wsAppVersion, tParam)
'                    If tResult(0).ToString = "0" Then
'                        Access_Token = tResult(2).ToString
'                    Else
'                        Access_Token = "x"
'                    End If

'                    ReDim CustomHeaders(0)
'                    CustomHeaders(0) = "X-API-Key|" & Access_Token

'                    Dim ParameterHeaders As String = ""
'                    For h As Integer = 0 To CustomHeaders.Length - 1
'                        If ParameterHeaders <> "" Then
'                            ParameterHeaders &= ","
'                        End If
'                        ParameterHeaders &= CustomHeaders(h)
'                    Next

'                    Dim mURLRCT As String = ""
'                    Try
'                        mURLRCT = ("" & ConfigurationManager.AppSettings("DeliveryManAPIUpdateStoreReceiptUrl")).Trim
'                    Catch ex As Exception
'                        mURLRCT = ""
'                    End Try
'                    If mURLOrder = "" Then
'                        mURLRCT = "x"
'                    End If


'                    For i As Integer = 0 To dtQuery.Rows.Count - 1

'                        Dim TrackId As String = dtQuery.Rows(i).Item("TrackId").ToString
'                        Dim TrackNum As String = dtQuery.Rows(i).Item("TrackNum").ToString
'                        Dim dOrderNo As String = dtQuery.Rows(i).Item("OrderNo").ToString
'                        Dim TrackTime As String = dtQuery.Rows(i).Item("TrackTime").ToString
'                        Dim WithError As String = "0" 'untuk penanda bila ada kegagalan

'                        Dim ObjRequest As New cPushTrackingRCTReq
'                        ObjRequest.orderId = dOrderNo
'                        ObjRequest.isStoreReceipt = True
'                        ObjRequest.receiptTimestamp = ConvertTimeUtc(TrackTime)

'                        Dim Parameter As String = JsonConvert.SerializeObject(ObjRequest)

'                        e.APIRequestLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, "Header:" & ParameterHeaders & "-Body:" & Parameter, mURLRCT, TrackNum & " " & dOrderNo)

'                        Dim Response As String = ObjFungsi.SendHTTP("", "", "", mURLRCT, Parameter, "", Encoding.Default, "60", "", mContentType, True, CustomHeaders, "PUT")
'                        Response = ("" & Response).Trim

'                        If Response <> "" Then

'                            e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, Response, mURLRCT, TrackNum & " " & dOrderNo)

'                            Dim ObjResponse As cPushTrackingRCTResp
'                            Try
'                                ObjResponse = JsonConvert.DeserializeObject(Of cPushTrackingRCTResp)(Response)

'                                If ObjResponse.status = "00" Then
'                                    'berhasil
'                                Else
'                                    WithError = "8"
'                                End If
'                            Catch ex As Exception
'                                WithError = "9"
'                                e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, ex, TrackNum & " " & dOrderNo)
'                            End Try

'                        Else
'                            WithError = "9"
'                            e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, "No Response", mURLRCT, TrackNum & " " & dOrderNo)

'                        End If 'dari If Response <> ""


'                        'apapun hasil-nya, anggap berhasil
'                        Query = "Update PushTracking"
'                        Query &= " Set `PushStatus` = '1'"
'                        If WithError <> "0" Then
'                            Query &= " , `RetryPush` = @WithError"
'                        End If
'                        Query &= " , UpdTime = now(), UpdUser = @AppName"
'                        Query &= " Where TrackId = @TrackId"
'                        Query &= " And PushStatus = '0'"

'                        SqlParam = New Dictionary(Of String, String)
'                        SqlParam.Add("@AppName", Strings.Left(wsAppName, 20))
'                        SqlParam.Add("@TrackId", TrackId)
'                        If WithError <> "0" Then
'                            SqlParam.Add("@WithError", WithError)
'                        End If

'                        ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

'                    Next ' dari For i As Integer = 0 To dtQuery.Rows.Count - 1

'                End If
'            End If 'dari If Not dtQuery Is Nothing

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, "FINISH")

'            Result = True

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, ex, Query)

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    <WebMethod()>
'    Public Function sendTrackEDO(ByVal AppName As String, ByVal AppVersion As String) As Boolean

'        Return PushTrackingEDO(AppName, AppVersion)

'    End Function

'    Private Function PushTrackingEDO(Optional ByVal AppName As String = "", Optional ByVal AppVersion As String = "") As Boolean

'        Dim Method As String = "PushTrackingEDO"

'        Dim Result As Boolean = False

'        Dim MCon As New MySqlConnection
'        Dim Query As String = ""

'        Try
'            MCon = MasterMCon.Clone

'            Dim e As New ClsError
'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, "START")


'            Dim nLimit As Integer = 0
'            Try
'                nLimit = CInt("" & ConfigurationManager.AppSettings("DeliveryManEditOrderLimit"))
'                If nLimit.ToString.Trim = "" Then
'                    nLimit = 0
'                End If
'                If IsNumeric(nLimit) = False Then
'                    nLimit = 0
'                End If
'            Catch ex As Exception
'                nLimit = 0
'            End Try
'            If nLimit < 1 Then
'                nLimit = 60
'            End If


'            Query = "Select TrackId, TrackNum, OrderNo, cast(TrackTime as char) as TrackTime"
'            Query &= " From PushTracking"
'            Query &= " Where Account = 'DMS' And TrackStatus = 'EDO' And PushStatus = '0'"
'            Query &= " Order By RetryPush, TrackTime"
'            Query &= " Limit " & nLimit

'            Dim ObjSQL As New ClsSQL
'            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)
'            If Not dtQuery Is Nothing Then
'                If dtQuery.Rows.Count > 0 Then

'                    For i As Integer = 0 To dtQuery.Rows.Count - 1

'                        Dim TrackId As String = dtQuery.Rows(i).Item("TrackId").ToString
'                        Dim TrackNum As String = dtQuery.Rows(i).Item("TrackNum").ToString
'                        Dim DeliveryId As String = dtQuery.Rows(i).Item("OrderNo").ToString
'                        Dim TrackTime As String = dtQuery.Rows(i).Item("TrackTime").ToString
'                        Dim WithError As String = "0" 'untuk penanda bila ada kegagalan

'                        Query = "Select d.* From TransactionDeliveryInfo d"
'                        Query &= " Inner Join `AutoOrderThirdPartyTransaction` a on (a.TrackNum = d.TrackNum)"
'                        Query &= " Where a.TrackNum = @TrackNum and a.dOrderNo = @DeliveryId"

'                        SqlParam = New Dictionary(Of String, String)
'                        SqlParam.Add("@TrackNum", TrackNum)
'                        SqlParam.Add("@DeliveryId", DeliveryId)

'                        Dim dtDeliveryInfo As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)


'                        '=== ikut pola Service.UpdateTransactionDeliveryInfo - BEGIN
'                        Try
'                            'susun datatable
'                            Dim dtService As New DataTable
'                            dtService.TableName = "SERVICE"
'                            dtService.Columns.Add("TrackNum")
'                            dtService.Columns.Add("OrderNo")
'                            dtService.Columns.Add("DeliveryId")
'                            dtService.Columns.Add("FlgBulky")
'                            dtService.Columns.Add("PinReturnBulky")

'                            'TrackNum, OrderNo, OrderType, OrderTypeDetail, ExpressType, FlgCOD, CODValue, CODValueDraft, CODPaymentCode, CODPaymentBiller
'                            ', FlgBulky, ReturnBulkyItems, ReturnBulkyItemsDraft, Notes
'                            ', ExpectedDeliverMinTime, ExpectedDeliverMinTimeDraft, ExpectedDeliverMaxTime, ExpectedDeliverMaxTimeDraft
'                            ', OrderPaidTime, ReceiptPrintTime, TimeZone, BasePIN, BasePINDraft, PINPickup, PINPickupDraft, PINCancel, PINCancelDraft
'                            ', PINKeep, PINKeepDraft, PINReturn, PINReturnDraft, AddInfo, AddTime, AddUser, UpdTime, UpdUser

'                            dtService.Rows.Add()
'                            dtService.Rows(dtService.Rows.Count - 1).Item("TrackNum") = TrackNum 'berdasarkan AWB Indopaket
'                            dtService.Rows(dtService.Rows.Count - 1).Item("OrderNo") = dtDeliveryInfo.Rows(0).Item("OrderNo").ToString 'nomor order partner
'                            dtService.Rows(dtService.Rows.Count - 1).Item("DeliveryId") = DeliveryId 'DeliveryId 3rd party
'                            dtService.Rows(dtService.Rows.Count - 1).Item("FlgBulky") = CInt(dtDeliveryInfo.Rows(0).Item("FlgBulky"))


'                            dtService.Rows(dtService.Rows.Count - 1).Item("PinReturnBulky") = ""
'                            If dtDeliveryInfo.Rows(0).Item("ReturnBulkyItems").ToString.Trim <> "" Then
'                                dtService.Rows(dtService.Rows.Count - 1).Item("PinReturnBulky") = dtDeliveryInfo.Rows(0).Item("PINReturn").ToString
'                            End If


'                            Dim dtCOD As New DataTable
'                            dtCOD.TableName = "COD"
'                            dtCOD.Columns.Add("CODValue")
'                            dtCOD.Columns.Add("CODValueDraft")

'                            dtCOD.Rows.Add()
'                            dtCOD.Rows(dtCOD.Rows.Count - 1).Item("CODValue") = CDbl(dtDeliveryInfo.Rows(0).Item("CODValue"))
'                            dtCOD.Rows(dtCOD.Rows.Count - 1).Item("CODValueDraft") = CDbl(dtDeliveryInfo.Rows(0).Item("CODValueDraft"))


'                            'informasi pengembalian barang bulky (galon atau lpg)
'                            Dim dtReturnBulky As New DataTable
'                            dtReturnBulky.TableName = "RETURNBULKY"
'                            dtReturnBulky.Columns.Add("Code")
'                            dtReturnBulky.Columns.Add("Description")
'                            dtReturnBulky.Columns.Add("Qty")
'                            Try
'                                If dtDeliveryInfo.Rows(0).Item("ReturnBulkyItems").ToString.Trim <> "" Then
'                                    Dim ReturnBulkyItems As String = dtDeliveryInfo.Rows(0).Item("ReturnBulkyItems").ToString.Trim
'                                    'Code1|Desc1=Qty1;Code2|Desc2=Qty2;...

'                                    Dim ReturnBulkyItemsSplit As String() = ReturnBulkyItems.Split(";")

'                                    For b As Integer = 0 To ReturnBulkyItemsSplit.Length - 1
'                                        If ReturnBulkyItemsSplit(b) <> "" Then
'                                            Dim ReturnBulkyItemsSplit2 As String() = ReturnBulkyItemsSplit(b).Split("|")
'                                            Dim ReturnBulkyItemsSplit3 As String() = ReturnBulkyItemsSplit2(1).Split("=")

'                                            Dim ReturnBulkyItemsCode As String = ReturnBulkyItemsSplit2(0)
'                                            Dim ReturnBulkyItemsDesc As String = ReturnBulkyItemsSplit3(0)
'                                            Dim ReturnBulkyItemsQty As Integer = CInt(ReturnBulkyItemsSplit3(1))

'                                            If ReturnBulkyItemsQty > 0 Then

'                                                dtReturnBulky.Rows.Add()
'                                                dtReturnBulky.Rows(dtReturnBulky.Rows.Count - 1).Item("Code") = ReturnBulkyItemsCode
'                                                dtReturnBulky.Rows(dtReturnBulky.Rows.Count - 1).Item("Description") = ReturnBulkyItemsDesc
'                                                dtReturnBulky.Rows(dtReturnBulky.Rows.Count - 1).Item("Qty") = ReturnBulkyItemsQty

'                                            End If
'                                        End If
'                                    Next
'                                End If
'                            Catch ex As Exception
'                            End Try


'                            'susun dataset
'                            Dim dsEditOrder As New DataSet
'                            dsEditOrder.Tables.Add(dtService)
'                            dsEditOrder.Tables.Add(dtCOD)
'                            dsEditOrder.Tables.Add(dtReturnBulky)

'                            Dim bParam(1) As Object
'                            bParam(0) = "system"
'                            bParam(1) = "#sY5t3M"

'                            Dim bResult As Object() = EditOrder(AppName, AppVersion, bParam, dsEditOrder)
'                            If bResult(0).ToString = "0" Then

'                            Else
'                                WithError = "9"
'                                e.DebugLog(MCon, AppName, AppVersion, Method, "", "Gagal EditOrder Deliveryman", TrackNum & " " & DeliveryId)
'                            End If


'                        Catch ex As Exception
'                            WithError = "9"
'                            e.ErrorLog(wsAppName, wsAppVersion, Method, wsAppName, ex, TrackNum & " " & DeliveryId)
'                        End Try
'                        '=== ikut pola Service.UpdateTransactionDeliveryInfo - END


'                        'apapun hasil-nya, anggap berhasil
'                        Query = "Update PushTracking"
'                        Query &= " Set `PushStatus` = '1'"
'                        If WithError <> "0" Then
'                            Query &= " , `RetryPush` = @WithError"
'                        End If
'                        Query &= " , UpdTime = now(), UpdUser = @AppName"
'                        Query &= " Where TrackId = @TrackId"
'                        Query &= " And PushStatus = '0'"

'                        SqlParam = New Dictionary(Of String, String)
'                        SqlParam.Add("@AppName", Strings.Left(wsAppName, 20))
'                        SqlParam.Add("@TrackId", TrackId)
'                        If WithError <> "0" Then
'                            SqlParam.Add("@WithError", WithError)
'                        End If

'                        ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

'                    Next ' dari For i As Integer = 0 To dtQuery.Rows.Count - 1

'                End If
'            End If 'dari If Not dtQuery Is Nothing

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, "FINISH")

'            Result = True

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(wsAppName, wsAppVersion, Method, wsAppName, ex, Query)

'        Finally
'            If MCon.State <> ConnectionState.Closed Then
'                MCon.Close()
'            End If
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Public Class cDoorPickupFinishStatusReq
'        Public orderId As String = "" 'delivery third party
'        Public delimanId As String = "" 'driver
'        Public delimanName As String = ""
'        Public delimanPhone As String = ""
'        Public delimanVehicle As String = ""
'        Public delimanCompany As String = ""
'    End Class

'    <WebMethod()>
'    Public Function DoorPickupFinishStatus(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        'cek status awb ipp, layanan door-to, apakah sudah DRO di toko asal

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "DoorPickupFinishStatus"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim Keyword As String = ""

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString

'            Dim JsonRequest As String = Param(2).ToString.Trim

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If JsonRequest = "" Then
'                Result(1) = "Input Request"
'                GoTo Skip
'            End If

'            Dim DeliveryId As String = ""

'            Dim ObjRequest As cDoorPickupFinishStatusReq
'            Try
'                ObjRequest = JsonConvert.DeserializeObject(Of cDoorPickupFinishStatusReq)(JsonRequest)
'                DeliveryId = ("" & ObjRequest.orderId.Trim)
'            Catch ex As Exception
'                DeliveryId = ""
'            End Try

'            If DeliveryId = "" Then
'                Result(1) = "Gagal generate DeliveryId"
'                GoTo Skip
'            End If

'            Keyword = DeliveryId

'            e.RequestLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "MULAI", Keyword)

'            Dim ObjSQL As New ClsSQL


'            Query = "Select Company1 as dTrackNum From AutoOrderTracking"
'            Query &= " Where `Status` = 'NEW' and Tracknum = @DeliveryId"

'            SqlParam = New Dictionary(Of String, String)
'            SqlParam.Add("@DeliveryId", DeliveryId)

'            Dim dTrackNum As String = ""
'            Try
'                dTrackNum = ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam).ToString.Trim.ToUpper
'            Catch ex As Exception
'                dTrackNum = ""
'            End Try

'            If dTrackNum = "" Then
'                Result(1) = DeliveryId & " tidak ditemukan"
'                GoTo Skip
'            End If


'            Query = "Select d.TrackNum From TransactionDoorPickupInfo d Where d.dTrackNum = @dTrackNum"

'            SqlParam = New Dictionary(Of String, String)
'            SqlParam.Add("@dTrackNum", dTrackNum)

'            Dim TrackNum As String = ("" & ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam)).ToString.Trim
'            If TrackNum = "" Or TrackNum = "-1" Then
'                Result(1) = "AWB Indopaket tidak ditemukan (" & DeliveryId & ")"
'                GoTo Skip
'            End If


'            Query = "Select r.TrackUserId as StoreCode, cast(r.TrackTime as char) as TrackTime"
'            Query &= " From Tracking r Where r.TrackNum = @TrackNum and r.`Status` = 'DRO'"

'            SqlParam = New Dictionary(Of String, String)
'            SqlParam.Add("@TrackNum", TrackNum)

'            Dim dtDRO As DataTable = Nothing
'            Try
'                dtDRO = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)
'            Catch ex As Exception
'                dtDRO = Nothing
'            End Try

'            If dtDRO Is Nothing Then
'                dtDRO = New DataTable
'            End If

'            If dtDRO.Rows.Count < 1 Then
'                Result(1) = "AWB " & TrackNum & " Belum Proses TPK (" & DeliveryId & ")"
'                GoTo Skip
'            End If


'            Result(0) = "0"
'            Result(1) = ""
'            Result(2) = dtDRO.Rows(0).Item("StoreCode").ToString.Trim.ToUpper & "#" & dtDRO.Rows(0).Item("TrackTime").ToString.Trim

'Skip:

'            Dim StrResult As String = Result(0).ToString
'            Try
'                StrResult &= "# " & Result(1).ToString
'            Catch ex As Exception
'            End Try
'            Try
'                StrResult &= "# " & Result(2).ToString
'            Catch ex As Exception
'            End Try

'            e.ResponseLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "SELESAI " & StrResult, Keyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, Query, Keyword)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    'disamakan dengan ServiceKlikIndogrosir
'    Public Class cHandoverScanQrCode
'        Public processType As String = "" 'PICKUP / CANCEL

'        Public kode_igr As String = ""
'        Public no_pb As String = "" 'orderno Igr
'        Public tgl_trans As String = ""
'        Public no_awb As String = "" 'awb IPP

'        Public orderId As String = "" 'nomor DMS
'        Public delimanId As String = ""
'        Public delimanName As String = ""
'        Public delimanPhone As String = ""
'        Public delimanVehicle As String = ""
'        Public delimanCompany As String = ""
'    End Class

'    <WebMethod()>
'    Public Function HandoverScanQrCode(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "HandoverScanQrCode"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim Keyword As String = ""

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString

'            Dim JsonRequest As String = Param(2).ToString.Trim

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If JsonRequest = "" Then
'                Result(1) = "Input Request"
'                GoTo Skip
'            End If

'            Dim ObjRequest As cHandoverScanQrCode
'            Try
'                ObjRequest = JsonConvert.DeserializeObject(Of cHandoverScanQrCode)(JsonRequest)
'            Catch ex As Exception
'                ObjRequest = Nothing
'            End Try

'            If ObjRequest Is Nothing Then
'                Result(1) = "Gagal generate Request"
'                GoTo Skip
'            End If

'            Dim ProcessType As String = ""
'            Try
'                ProcessType = ("" & ObjRequest.processType.Trim.ToUpper)
'            Catch ex As Exception
'                ProcessType = ""
'            End Try
'            If ProcessType = "" Then
'                ProcessType = "PICKUP" 'sementara hanya ada 1 jenis
'                'Result(1) = "Gagal get ProcessType"
'                'GoTo Skip
'            End If

'            Dim TrackStatus As String = ""
'            Select Case ProcessType
'                Case "PICKUP"
'                    TrackStatus = "AQR"
'                Case "CANCEL"
'                    TrackStatus = "BQR"
'                Case Else
'                    Result(1) = "ProcessType " & ProcessType & " tidak dikenali"
'                    GoTo Skip
'            End Select

'            Dim AwbIpp As String = ""
'            Try
'                AwbIpp = ("" & ObjRequest.no_awb.Trim)
'            Catch ex As Exception
'                AwbIpp = ""
'            End Try
'            If AwbIpp = "" Then
'                Result(1) = "Gagal get AWB IPP"
'                GoTo Skip
'            End If

'            Dim DmsId As String = ""
'            Try
'                DmsId = ("" & ObjRequest.orderId.Trim)
'            Catch ex As Exception
'                DmsId = ""
'            End Try
'            If DmsId = "" Then
'                Result(1) = "Gagal get nomor DMS"
'                GoTo Skip
'            End If

'            Dim Kode_Igr As String = ""
'            Try
'                Kode_Igr = ("" & ObjRequest.kode_igr.Trim)
'            Catch ex As Exception
'                Kode_Igr = ""
'            End Try
'            If Kode_Igr = "" Then
'                Result(1) = "Gagal get Kode IGR"
'                GoTo Skip
'            End If

'            Dim No_Pb As String = ""
'            Try
'                No_Pb = ("" & ObjRequest.no_pb.Trim)
'            Catch ex As Exception
'                No_Pb = ""
'            End Try
'            If No_Pb = "" Then
'                Result(1) = "Gagal get nomor PB"
'                GoTo Skip
'            End If


'            Keyword = AwbIpp & " " & ProcessType

'            e.RequestLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "MULAI", Keyword)

'            Dim ObjSQL As New ClsSQL


'            Query = "Select a.Account, a.ThirdParty, t.CoName"
'            Query &= " From AutoOrderThirdPartyTransaction a"
'            Query &= " Inner Join `Transaction` t on (a.TrackNum = t.TrackNum)"
'            Query &= " Where a.Tracknum = @AwbIpp and a.dOrderNo = @DmsId"

'            SqlParam = New Dictionary(Of String, String)
'            SqlParam.Add("@AwbIpp", AwbIpp)
'            SqlParam.Add("@DmsId", DmsId)

'            Dim dtAutoOrder As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)

'            If dtAutoOrder Is Nothing Then
'                Result(1) = "Data is nothing"
'                GoTo Skip
'            End If

'            If dtAutoOrder.Rows.Count < 1 Then
'                Result(1) = "Data tidak ditemukan (" & DmsId & " " & AwbIpp & ")"
'                GoTo Skip
'            End If

'            Query = "Insert Into PushTracking ("
'            Query &= " Account, TrackNum, OrderNo, TrackStatus, TrackTime, TrackUser, TrackUserID, Opr1, OprName2, AddInfo, AddTime, AddUser"
'            Query &= " ) values ("
'            Query &= " @Account, @AwbIpp, @No_Pb, @TrackStatus, now(), 'SYS', @Kode_Igr, @ThirdParty, @CoName, @JsonRequest, now(), @AddUser"
'            Query &= " )"

'            SqlParam = New Dictionary(Of String, String)
'            SqlParam.Add("@Account", dtAutoOrder.Rows(0).Item("Account").ToString)
'            SqlParam.Add("@AwbIpp", AwbIpp)
'            SqlParam.Add("@No_Pb", No_Pb)
'            SqlParam.Add("@Kode_Igr", Kode_Igr)
'            SqlParam.Add("@TrackStatus", TrackStatus)
'            SqlParam.Add("@ThirdParty", dtAutoOrder.Rows(0).Item("ThirdParty").ToString)
'            SqlParam.Add("@CoName", dtAutoOrder.Rows(0).Item("CoName").ToString)
'            SqlParam.Add("@JsonRequest", JsonRequest)
'            SqlParam.Add("@AddUser", User)

'            If ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam) = False Then
'                e.DebugLog(MCon, AppName, AppVersion, Method, User, "Gagal Insert PushTracking " & TrackStatus, Keyword)
'            End If

'            Result(0) = "0"
'            Result(1) = ""
'            Result(2) = ""

'Skip:

'            e.ResponseLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "SELESAI", Keyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, Query, Keyword)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    <WebMethod()>
'    Public Function LogHeaderParam(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

'        'hanya untuk testing

'        Dim Result As Object = CreateResult()

'        Dim Method As String = "LogHeaderParam"

'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim LogKeyword As String = Format(Date.Now, "yyMMddHHmmss")

'        Dim User As String = Param(0).ToString

'        Try
'            Dim Password As String = Param(1).ToString

'            Dim Parameter As String = Param(2).ToString.Trim

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim Process As Boolean = False

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            If Parameter = "" Then
'                Result(1) = "Input parameter"
'                GoTo Skip
'            End If

'            Dim mURLLogHeaderParam As String = ("" & ConfigurationManager.AppSettings("CallbackLogHeaderParamUrl")).Trim
'            If mURLLogHeaderParam = "" Then
'                mURLLogHeaderParam = "x"
'            End If

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "MULAI " & Parameter, LogKeyword)

'            ReDim CustomHeaders(1)
'            CustomHeaders(0) = "x-api-key" & "|" & "budil"
'            CustomHeaders(1) = "UserAgent" & "|" & "INDOPAKET"

'            Dim ParameterHeaders As String = ""
'            For h As Integer = 0 To CustomHeaders.Length - 1
'                If ParameterHeaders <> "" Then
'                    ParameterHeaders &= ","
'                End If
'                ParameterHeaders &= CustomHeaders(h)
'            Next

'            Dim ObjFungsi As New ClsFungsi

'            e.APIRequestLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "Header:" & ParameterHeaders & "-" & "Body:" & Parameter, mURLLogHeaderParam, LogKeyword)

'            Dim Response As String = ""
'            Response = ObjFungsi.SendHTTP("", "", "", mURLLogHeaderParam, Parameter, "", Encoding.Default, mTimeout, "", mContentType, True, CustomHeaders)
'            Response = ("" & Response).Trim

'            If Response <> "" Then
'                e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, Response, mURLLogHeaderParam, LogKeyword)
'            Else
'                e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "No Response", mURLLogHeaderParam, LogKeyword)
'            End If 'dari If Response <> ""


'            e.APIRequestLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "Body:" & Parameter, mURLLogHeaderParam, LogKeyword)

'            Response = ObjFungsi.SendHTTP("", "", "", mURLLogHeaderParam, Parameter, "INDOPAKET", Encoding.Default, mTimeout, "", mContentType, True)
'            Response = ("" & Response).Trim

'            If Response <> "" Then
'                e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, Response, mURLLogHeaderParam, LogKeyword)
'            Else
'                e.APIResponseLog(MCon, wsAppName, wsAppVersion, "WServiceDeliMan." & Method, wsUser, "No Response", mURLLogHeaderParam, LogKeyword)
'            End If 'dari If Response <> ""


'            Result(0) = "0"
'            Result(1) = "OK"
'            Result(2) = ""

'            Process = True

'Skip:

'            e.DebugLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", "SELESAI " & Process.ToString, LogKeyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, "WServiceDeliMan", ex, Query)

'            Result(1) &= "Error : " & ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Public Class cSendDataKaryawanReq
'        Public delivery_man_id As String = "" 'nik karyawan
'        Public delivery_man_name As String = "" 'nama karyawan
'        Public prev_delivery_man_id As String = "" 'nik sebelumnya
'        Public effective_date As String = "" 'tanggal efektif pergantian nik YYYY-MM-DD
'    End Class

'    Public Class cSendDataKaryawanRsp
'        '00 = success
'        '03 = bad request

'        Public status As String = ""
'        Public message As String = ""
'        'Public data As object
'        Public timestamp As String = ""
'        Public signature As String = ""
'    End Class

'    <WebMethod()>
'    Public Function SendDataKaryawan(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object

'        Dim Method As String = "SendDataKaryawan"

'        Dim Result As Object() = CreateResult()
'        Dim MCon As New MySqlConnection

'        Dim Query As String = ""

'        Dim LogKeyword As String = Format(Date.Now, "yyMMddHHmmss")

'        Dim User As String = Param(0).ToString
'        Try
'            Dim Password As String = Param(1).ToString

'            MCon = MasterMCon.Clone
'            MCon.Open()

'            Dim e As New ClsError

'            Dim UserOK As Object() = WService.ValidasiUser(MCon, User, Password)
'            If UserOK(0) <> "0" Then
'                Result(1) = UserOK(1)
'                GoTo Skip
'            End If

'            e.DebugLog(MCon, AppName, AppVersion, Method, User, "START", LogKeyword)

'            Dim OrderLimit As Integer = 0
'            Try
'                Dim sOrderLimit As String = ("" & ConfigurationManager.AppSettings("DeliveryManSendDataKaryawanLimit")).Trim
'                If IsNumeric(sOrderLimit) Then
'                    OrderLimit = CInt(sOrderLimit)
'                Else
'                    OrderLimit = 0
'                End If
'            Catch ex As Exception
'                OrderLimit = 0
'            End Try
'            If OrderLimit < 1 Then
'                OrderLimit = 120
'            End If

'            '-- kodecabang = pushstatus; kodebagian = retry (sudah sebagai index juga)
'            Query = "Select *"
'            Query &= " From HrdDataKaryawan_PushDms"
'            Query &= " Where KodeCabang = '0'"
'            Query &= " Order By KodeBagian, UpdTime"
'            Query &= " Limit " & OrderLimit

'            Dim ObjSQL As New ClsSQL
'            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)
'            If Not dtQuery Is Nothing Then
'                If dtQuery.Rows.Count > 0 Then

'                    Dim ObjFungsi As New ClsFungsi

'                    Dim AlreadyResetToken As Boolean = False
'                    Dim MustResetToken As Boolean = False

'UlangGetOAuth:

'                    Dim Access_Token As String = ""

'                    Dim tParam() As Object
'                    If MustResetToken = False Then
'                        ReDim tParam(1)
'                        tParam(0) = Param(0)
'                        tParam(1) = Param(1)
'                    Else
'                        ReDim tParam(2)
'                        tParam(0) = Param(0)
'                        tParam(1) = Param(1)
'                        tParam(2) = "1"
'                        AlreadyResetToken = True
'                    End If

'                    Dim tResult As Object() = GetOAuth(wsAppName, wsAppVersion, tParam)
'                    If tResult(0).ToString = "0" Then
'                        Access_Token = tResult(2).ToString
'                    Else
'                        Result(1) = tResult(1).ToString
'                        GoTo Skip
'                    End If

'                    ReDim CustomHeaders(0)
'                    CustomHeaders(0) = "X-API-Key|" & Access_Token

'                    Dim ParameterHeaders As String = ""
'                    For h As Integer = 0 To CustomHeaders.Length - 1
'                        If ParameterHeaders <> "" Then
'                            ParameterHeaders &= ","
'                        End If
'                        ParameterHeaders &= CustomHeaders(h)
'                    Next

'                    For i As Integer = 0 To dtQuery.Rows.Count - 1

'                        Dim DelimanId As String = dtQuery.Rows(i).Item("NIK").ToString.Trim
'                        Dim DelimanName As String = dtQuery.Rows(i).Item("Nama").ToString.Trim
'                        Dim PrevDelimanId As String = dtQuery.Rows(i).Item("NIKLama").ToString.Trim
'                        Dim EffectiveDate As String = ""
'                        Try
'                            EffectiveDate = Format(CDate(dtQuery.Rows(i).Item("TglEfektifDHR").ToString.Trim), "yyyy-MM-dd")
'                        Catch ex As Exception
'                            EffectiveDate = ""
'                        End Try
'                        'If PrevDelimanId <> "" And EffectiveDate = "" Then
'                        '    EffectiveDate = Format(Date.Now, "yyyy-MM-dd")
'                        'End If

'                        Dim ObjRequest As New cSendDataKaryawanReq
'                        ObjRequest.delivery_man_id = DelimanId
'                        ObjRequest.delivery_man_name = DelimanName
'                        ObjRequest.prev_delivery_man_id = PrevDelimanId
'                        ObjRequest.effective_date = EffectiveDate

'                        Dim Parameter As String = JsonConvert.SerializeObject(ObjRequest)

'                        e.APIRequestLog(MCon, wsAppName, wsAppVersion, Method, wsUser, "Header:" & ParameterHeaders & "-" & "Body:" & Parameter, mURLSendDataKaryawan, DelimanId)

'                        Dim Response As String = ""
'                        Response = ObjFungsi.SendHTTP("", "", "", mURLSendDataKaryawan, Parameter, "", Encoding.Default, mTimeout, "", mContentType, True, CustomHeaders)
'                        Response = ("" & Response).Trim

'                        If Response <> "" Then
'                            e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsUser, Response, mURLSendDataKaryawan, DelimanId)
'                            Try
'                                Dim ObjResponse As cSendDataKaryawanRsp = JsonConvert.DeserializeObject(Of cSendDataKaryawanRsp)(Response)
'                                If ObjResponse.status = "00" Then
'                                    Query = "Update HrdDataKaryawan_PushDms"
'                                    Query &= " Set KodeCabang = '1', KodeBagian = '0'"
'                                    Query &= " , UpdTime = now(), UpdUser = @UpdUser"
'                                    Query &= " Where KodeCabang = '0' And NIK = @DelimanId"
'                                Else
'                                    Query = "Update HrdDataKaryawan_PushDms"
'                                    Query &= " Set KodeBagian = KodeBagian + 1"
'                                    Query &= " , UpdTime = now(), UpdUser = @UpdUser"
'                                    Query &= " Where KodeCabang = '0' And NIK = @DelimanId"
'                                End If

'                                SqlParam = New Dictionary(Of String, String)

'                                SqlParam.Add("@DelimanId", DelimanId)
'                                SqlParam.Add("@UpdUser", wsAppName & " " & wsAppVersion & " " & Method)

'                                ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

'                            Catch ex As Exception
'                                e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, wsUser, ex, "", LogKeyword)
'                            End Try

'                        Else
'                            e.APIResponseLog(MCon, wsAppName, wsAppVersion, Method, wsUser, "No Response", mURLSendDataKaryawan, DelimanId)

'                            Query = "Update HrdDataKaryawan_PushDms"
'                            Query &= " Set KodeBagian = KodeBagian + 1"
'                            Query &= " , UpdTime = now(), UpdUser = @UpdUser"
'                            Query &= " Where KodeCabang = '0' And NIK = @DelimanId"

'                            SqlParam = New Dictionary(Of String, String)

'                            SqlParam.Add("@DelimanId", DelimanId)
'                            SqlParam.Add("@UpdUser", wsAppName & " " & wsAppVersion & " " & Method)

'                            ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

'                        End If 'dari If Response <> ""

'                    Next

'                End If 'dari If dtQuery.Rows.Count > 0
'            End If 'dari If Not dtQuery Is Nothing

'            Result(2) = ""
'            Result(1) = ""
'            Result(0) = "0"

'Skip:

'            e.DebugLog(MCon, AppName, AppVersion, Method, User, "FINISH", LogKeyword)

'        Catch ex As Exception
'            Dim e As New ClsError
'            e.ErrorLog(MCon, AppName, AppVersion, Method, User, ex, "", LogKeyword)

'            Result(1) = ex.Message

'        Finally
'            Try
'                MCon.Close()
'            Catch ex As Exception
'            End Try
'            Try
'                MCon.Dispose()
'            Catch ex As Exception
'            End Try

'        End Try

'        Return Result

'    End Function


'    Dim wsAppName As String = "CorSvcDeliveryMan"
'    Dim wsAppVersion As String = "24.08.28.00"

'    <WebMethod()>
'    Public Function A_WSDevelopment() As String
'        Return "This is WS Development"
'    End Function

'End Class
