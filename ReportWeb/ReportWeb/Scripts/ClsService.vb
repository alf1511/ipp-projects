Imports Microsoft.VisualBasic
Imports System.Data
Imports System.IO
Imports ReportWeb.ClsWebVer

Public Class ClsService

    Dim serv As New CoreService.Service
    Dim servPPC As New ServiceProfilePic.ServiceProfilePic
    Dim param_() As Object
    Dim respon_() As Object

    Dim ObjFungsi As New ClsFungsi

    Public Function GetECommerceList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetECommerceList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetAccountDCList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountDCList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountDCList(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountDCList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountCityDoorList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountCityDoorList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountCityDoorList(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountCityDoorList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountCityDoorToPickupList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountCityDoorToPickupList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        End Try

    End Function

    Public Function SetAccountCityDoorToPickupList(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountCityDoorToPickupList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountCityPickupList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountCityPickupList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountCityPickupList(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountCityPickupList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountCityTruckingList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountCityTruckingList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountCityTruckingList(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountCityTruckingList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountCountryList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim dsData As New DataSet
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountCountryList(AppName, AppVersion, param, dsData)

            If respon_(0) = "0" Then

                dt = dsData.Tables(0).Copy
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountCountryList(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountCountryList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountFleetRentalVehicleTypeList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountFleetRentalVehicleTypeList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        End Try

    End Function

    Public Function SetAccountFleetRentalVehicleTypeList(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountFleetRentalVehicleTypeList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function PrintTrackNumUpload(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PrintTrackNumUpload(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountMaxWeight(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountMaxWeight(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountMaxWeight(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountMaxWeight(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountMaxWeight3PL(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountMaxWeight3PL(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountMaxWeight3PL(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountMaxWeight3PL(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountMaxPckValue(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountMaxPckValue(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountMaxPckValue(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountMaxPckValue(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountMaxAndSpcDimension(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountMaxAndSpcDimension(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountMaxAndSpcDimension(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountMaxAndSpcDimension(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountMaxCODValue(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountMaxCODValue(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountMaxCODValue(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountMaxCODValue(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountCoAddrNoNum(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountCoAddrNoNum(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetAccountCoAddrNoNum(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountCoAddrNoNum(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function AutoCopyAndReturnGetTrackNumList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.AutoCopyAndReturnGetTrackNumList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        End Try

    End Function

    Public Function GetCopyAndReturnTrackNumList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetCopyAndReturnTrackNumList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetCopyAndReturnAccountList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetCopyAndReturnAccountList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function OtherExpeditionExpenseData(ByVal param() As Object, ByRef dtHub As DataTable, ByRef dtKota As DataTable, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.OtherExpeditionExpenseData(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                dtHub = ObjFungsi.ConvertStringToDatatable(respon_(3))
                dtKota = ObjFungsi.ConvertStringToDatatable(respon_(4))

                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function OtherExpeditionExpenseUpdate(ByVal param() As Object, ByRef ErrMsg As String) As Boolean

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.OtherExpeditionExpenseUpdate(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                ErrMsg = respon_(1)
                Return True

            Else

                ErrMsg = respon_(1)
                Return False

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return False

        End Try

    End Function

    Public Function OtherExpeditionExpenseDelete(ByVal param() As Object, ByRef ErrMsg As String) As Boolean

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.OtherExpeditionExpenseDelete(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                ErrMsg = respon_(1)
                Return True

            Else

                ErrMsg = respon_(1)
                Return False

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return False

        End Try

    End Function

    Public Function CheckTransaction(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.CheckTransaction(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function TrackingUpdateStatus(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.TrackingUpdateStatus(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function PrintSuratJalan(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PrintSuratJalan(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function MotherstoreValidation(ByVal param() As Object, ByRef dtKodePOS As DataTable, ByRef dtMotherStore As DataTable, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.MotherstoreValidation(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                dtKodePOS = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(3))
                dtMotherStore = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(4))

                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function PettyCashHubExpenseAdd(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PettyCashHubExpenseAdd(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function PettyCashHubExpenseDelete(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PettyCashHubExpenseDelete(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function PettyCashHubExpenseDeleteRecord(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PettyCashHubExpenseDeleteRecord(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function PettyCashHubFinalizeRecord(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PettyCashHubFinalizeRecord(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function PettyCashHubExpenseApproval(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PettyCashHubExpenseApproval(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function PettyCashHubProcessJournal(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PettyCashHubProcessJournal(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function PettyCashReportList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PettyCashReportList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function PettyCashSubCategoryList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PettyCashSubCategoryList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function PettyCashHubFindRecord(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.PettyCashHubFindRecord(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function HubAccountPickupListGet(ByVal param() As Object, ByVal Hub As String, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.HubAccountPickupListGet(AppName, AppVersion, Hub, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function HubAccountPickupListSet(ByVal param() As Object, ByVal Hub As String, ByRef ErrMsg As String) As Object

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.HubAccountPickupListSet(AppName, AppVersion, Hub, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function DashboardHubAccountListGet(ByVal param() As Object, ByVal Hub As String, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.DashboardHubAccountListGet(AppName, AppVersion, Hub, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function DashboardHubAccountListSet(ByVal param() As Object, ByVal Hub As String, ByRef ErrMsg As String) As Object

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.DashboardHubAccountListSet(AppName, AppVersion, Hub, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function FunctionCreateFile(ByVal FileSource As String, ByVal FileDestination As String, ByRef ErrorMessage As String) As Boolean

        Try

            Dim Param(1) As Object
            Param(0) = FileSource
            Param(1) = FileDestination

            ObjFungsi.WriteTracelogTxt("Debug FunctionCreateFile, FileSource : " & FileSource & ", FileDestination : " & FileDestination & " start")

            Dim respon() As Object = servPPC.FunctionCreateFile(AppName, AppVersion, Param)

            ObjFungsi.WriteTracelogTxt("Debug FunctionCreateFile FileSource : " & FileSource & ", FileDestination : " & FileDestination & ", Respon(0) : " & respon(0) & ", Respon(1) : " & respon(1))

            If respon(0) = 0 Then
                Return True
            Else
                ErrorMessage = respon(1)
                Return False
            End If

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return False

        Finally

            ObjFungsi.WriteTracelogTxt("Debug FunctionCreateFile FileSource : " & FileSource & ", FileDestination : " & FileDestination & ", finish")

        End Try

    End Function

    Public Function FunctionDeleteFile(ByVal FileToDelete As String, ByRef ErrorMessage As String) As Boolean

        Try

            Dim Param(0) As Object
            Param(0) = FileToDelete

            ObjFungsi.WriteTracelogTxt("Debug FunctionDeleteFile, FileToDelete : " & FileToDelete & " start")

            Dim respon() As Object = servPPC.FunctionDeleteFile(AppName, AppVersion, Param)

            ObjFungsi.WriteTracelogTxt("Debug FunctionDeleteFile FileToDelete : " & FileToDelete & ", Respon(0) : " & respon(0) & ", Respon(1) : " & respon(1))

            If respon(0) = 0 Then
                Return True
            Else
                ErrorMessage = respon(1)
                Return False
            End If

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return False

        Finally

            ObjFungsi.WriteTracelogTxt("Debug FunctionDeleteFile FileToDelete : " & FileToDelete & ", finish")

        End Try

    End Function

    Public Function FunctionGetFileInfoList(ByVal Directory As String, ByRef dsData As DataSet, ByRef ErrorMessage As String) As Boolean

        Dim dirInfo As FileInfo() = Nothing

        Try

            Dim Param(0) As Object
            Param(0) = Directory

            ObjFungsi.WriteTracelogTxt("Debug FunctionGetFileInfoList, Directory : " & Directory & " start")

            Dim respon() As Object = servPPC.FunctionGetFileInfoList(AppName, AppVersion, Param, dsData)

            ObjFungsi.WriteTracelogTxt("Debug FunctionGetFileInfoList FileSource : " & Directory & ", ErrorMessage : " & ErrorMessage)

            If respon(0) = 0 Then
                Return True
            Else
                ErrorMessage = respon(1)
                Return False
            End If

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return False

        Finally

            ObjFungsi.WriteTracelogTxt("Debug FunctionGetFileInfoList Directory : " & Directory & ", finish")

        End Try

    End Function

    Public Function AWBTTFSupplierGetList(ByRef DataAWB1 As DataTable, ByRef DataAWB2 As DataTable) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = UserWS
            param(1) = PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.AWBTTFSupplierGetList(AppName, AppVersion, param)

            If respon(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

                DataAWB1 = ObjFungsi.ConvertStringToDatatableWithBarcode(respon(3))
                DataAWB2 = ObjFungsi.ConvertStringToDatatableWithBarcode(respon(4))

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

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

    Public Function AWBTTFSupplierCreateAWB(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.AWBTTFSupplierCreateAWB(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function SuratTugasKurirIPPGetStore(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SuratTugasKurirIPPGetStore(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SuratTugasKurirIPPGetAWBList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SuratTugasKurirIPPGetAWBList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SuratTugasKurirIPPProcess(ByVal param() As Object, ByVal Timeout As Integer, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            serv.Timeout = Timeout
            respon_ = serv.SuratTugasKurirIPPProcess(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function DocumentUploadCategoryList(ByVal ViewType As String, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim param(2) As Object
            param(0) = UserWS
            param(1) = PassWS
            param(2) = ViewType

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.DocumentUploadCategoryList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function DocumentUploadFindRecord(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.DocumentUploadFindRecord(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function DocumentUploadAddFile(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.DocumentUploadAddFile(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function TanyaDataITTokoCategoryList(ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim param(1) As Object
            param(0) = UserWS
            param(1) = PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.TanyaDataITTokoCategoryList(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function TanyaDataITTokoGetData(ByVal param() As Object, ByRef StoreInfo As String, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.TanyaDataITTokoGetData(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                StoreInfo = respon_(1)
                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function TanyaDataITTokoEmail(ByVal param() As Object, ByRef ErrMsg As String) As Object

        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.TanyaDataITTokoEmail(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetAccountSetting(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing
        Dim DsData As New DataSet

        Try
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountSetting(AppName, AppVersion, param, DsData)

            If respon_(0) = "0" Then

                dt = DsData.Tables(0).Copy
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        End Try

    End Function

    Public Function GetAccountSettingTypeList(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing
        Dim dsData As New DataSet
        Try

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetAccountSettingTypeList(AppName, AppVersion, param, dsData)

            If respon_(0) = "0" Then

                dt = dsData.Tables(0).Copy
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        End Try
    End Function

    Public Function SetAccountSetting(ByVal param() As Object, ByRef DsData As DataSet, ByRef ErrMsg As String) As Object

        Try
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.SetAccountSetting(AppName, AppVersion, param, DsData)

            If respon_(0) = "0" Then

                Return respon_

            Else

                ErrMsg = respon_(1)
                Return respon_

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return Nothing

        End Try

    End Function

    Public Function GetDoorPickupSlot(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.GetDoorPickupSlot(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        End Try

    End Function

    Public Function DashboardPartnerPickupGetRequestAddress(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.DashboardPartnerPickupGetRequestAddress(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        End Try

    End Function

    Public Function DashboardPartnerPickupPairAWBPJB(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.DashboardPartnerPickupPairAWBPJB(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        End Try

    End Function

    Public Function DashboardPartnerPickupGetDetailReqId(ByVal param() As Object, ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon_ = serv.DashboardPartnerPickupGetDetailReqId(AppName, AppVersion, param)

            If respon_(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))
                Return dt

            Else

                ErrMsg = respon_(1)
                Return dt

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsService, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)
            Return dt

        End Try

    End Function

    Public Function DashboardPartnerPickupGetPickupRequestTemplateList() As DataTable

        Dim dt As New DataTable

        ReDim param_(1)
        param_(0) = UserWS
        param_(1) = PassWS

        respon_ = Nothing

        Dim i As Integer = 1
        Dim success As Boolean = False
        While i <= maxTryWS And success = False
            Try
                respon_ = serv.DashboardPartnerPickupGetPickupRequestTemplateList(AppName, AppVersion, param_)
                success = True

            Catch ex As Exception

                If i >= maxTryWS Then
                    Throw
                End If

                i = i + 1

            End Try

        End While

        If respon_(0) = "0" Then

            dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))

        Else

            Return Nothing

        End If

        Return dt

    End Function
    Public Function GetTemplatePickupRequestAddress(ByRef Pesan As String, ByVal Template As String, ByVal Type As String) As DataTable

        Pesan = ""

        Dim dt As New DataTable

        ReDim param_(3)
        param_(0) = UserWS
        param_(1) = PassWS
        param_(2) = Template
        param_(3) = Type

        respon_ = Nothing

        Dim i As Integer = 1
        Dim success As Boolean = False
        While i <= maxTryWS And success = False

            Try

                respon_ = serv.GetTemplatePickupRequestAddress(AppName, AppVersion, param_)
                success = True

            Catch ex As Exception

                If i >= maxTryWS Then
                    Throw
                End If

                i = i + 1

            End Try

        End While

        If respon_(0) = "0" Then

            dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))

        Else

            Pesan = "Gagal mengambil data !!"

            Return Nothing

        End If

        Return dt

    End Function

    Public Function GetAutoPickupRequestSetting(ByRef Pesan As String, ByVal Template As String) As DataTable

        Pesan = ""

        Dim dt As New DataTable

        ReDim param_(2)
        param_(0) = UserWS
        param_(1) = PassWS
        param_(2) = Template

        respon_ = Nothing

        Dim i As Integer = 1
        Dim success As Boolean = False
        While i <= maxTryWS And success = False

            Try

                respon_ = serv.GetAutoPickupRequestSetting(AppName, AppVersion, param_)
                success = True

            Catch ex As Exception

                If i >= maxTryWS Then
                    Throw
                End If

                i = i + 1

            End Try

        End While

        If respon_(0) = "0" Then

            dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))

        Else

            Pesan = "Gagal mengambil data !!"

            Return Nothing

        End If

        Return dt

    End Function

    Public Function DashboardPartnerAccountList() As DataTable

        Dim dt As New DataTable

        ReDim param_(4)
        param_(0) = UserWS
        param_(1) = PassWS
        param_(4) = "REQUESTPICKUPIPPHO"

        respon_ = Nothing

        Dim i As Integer = 1
        Dim success As Boolean = False
        While i <= maxTryWS And success = False

            Try
                respon_ = serv.GetECommerceList(AppName, AppVersion, param_)
                success = True

            Catch ex As Exception

                If i >= maxTryWS Then
                    Throw
                End If

                i = i + 1

            End Try

        End While

        If respon_(0) = "0" Then

            dt = ObjFungsi.ConvertStringToDatatable(respon_(2))

        Else

            Return Nothing

        End If

        Return dt

    End Function

    Public Function DashboardPartnerPickupGetSlot() As DataTable

        Dim dt As New DataTable

        ReDim param_(1)
        param_(0) = UserWS
        param_(1) = PassWS

        respon_ = Nothing

        Dim i As Integer = 1
        Dim success As Boolean = False
        While i <= maxTryWS And success = False

            Try
                respon_ = serv.DashboardPartnerPickupGetSlot(AppName, AppVersion, param_)
                success = True

            Catch ex As Exception

                If i >= maxTryWS Then
                    Throw
                End If

                i = i + 1

            End Try

        End While

        If respon_(0) = "0" Then

            dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))

        Else

            Return Nothing

        End If

        Return dt

    End Function

    Public Function GetFavoriteAddressListIPPHO(ByRef Pesan As String, ByVal Username As String, ByVal Account As String) As DataTable

        Pesan = ""

        Dim dt As New DataTable

        ReDim param_(3)
        param_(0) = UserWS
        param_(1) = PassWS
        param_(2) = Username
        param_(3) = Account

        respon_ = Nothing

        Dim i As Integer = 1
        Dim success As Boolean = False
        While i <= maxTryWS And success = False

            Try

                respon_ = serv.GetFavoriteAddressListIPPHO(AppName, AppVersion, param_)
                success = True

            Catch ex As Exception

                If i >= maxTryWS Then
                    Throw
                End If

                i = i + 1

            End Try

        End While

        If respon_(0) = "0" Then

            dt = ObjFungsi.ConvertStringToDatatableWithBarcode(respon_(2))

        Else

            Pesan = "Gagal mengambil data !!"

            Return Nothing

        End If

        Return dt

    End Function

End Class
