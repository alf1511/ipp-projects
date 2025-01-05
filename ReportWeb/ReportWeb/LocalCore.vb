Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports MySql.Data.MySqlClient
Imports Newtonsoft.Json
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared
Imports System.Diagnostics
Imports Barcode
Imports System.Net.Mail
Imports Pop3
Imports System.Text.ASCIIEncoding
Imports System.IO
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Net
Imports Microsoft.VisualBasic.ApplicationServices
Imports Org.BouncyCastle.Crypto.Engines
Imports Google.Protobuf.WellKnownTypes
Imports System.Web.UI.WebControls.Expressions



Public Class LocalCore
    Private SqlParam As New Dictionary(Of String, String)
    Private ResponseOK As String = "BERHASIL"
    Private MasterMCon As MySqlConnection
    Private Master As MySqlConnection
    Private Master2MCon As MySqlConnection
    Private MasterMConLive As MySqlConnection
    Private MasterMConSlave1 As MySqlConnection
    Private MasterMConSlave2 As MySqlConnection
    Private wsAppName As String = "CoreService"
    Private wsAppVersion As String = "24.09.12.00"
    Private StringTest As String = "TEST"
    Private CustomAccountKlikApka As String
    Private CustomAccountKlikIStore As String
    Private CustomAccountKlikFood As String
    Private CustomAccountKlikReOrder As String
    Private OtherExpeditionDefault As String = ".3000000.3000013." '3PL yang tidak perlu input nomor resi, exclude mobil konsol (aktif live 230411)


    Private GlobalSjNumShortLength As Integer = 7
    Private MyDbName As String = "iexpress_development"

    Public Sub New()
        Dim ConSQL As String = "" & ConfigurationManager.AppSettings("ConSQL")

        If ConSQL <> "" Then
            'webconfig
            MasterMCon = New MySqlConnection(ConSQL)
        Else
            'default
            MasterMCon = New MySqlConnection("Server=db-development.cwaihqqxerhv.ap-southeast-1.rds.amazonaws.com;Port=3306;Database=dev_jenius;Uid=dev_indopaket;Pwd=pindopaket;")
        End If

        Dim Con2SQL As String = "" & ConfigurationManager.AppSettings("Con2SQL")

        If Con2SQL <> "" Then
            'webconfig
            Master2MCon = New MySqlConnection(Con2SQL)
        Else
            'default
            Master2MCon = New MySqlConnection("Server=db-development.cwaihqqxerhv.ap-southeast-1.rds.amazonaws.com;Port=3306;Database=dev_jenius;Uid=dev_indopaket;Pwd=pindopaket;")
        End If

        Dim ConSQLLive As String = "" & ConfigurationManager.AppSettings("ConSQLLive")

        If ConSQLLive <> "" Then
            'webconfig
            MasterMConLive = New MySqlConnection(ConSQLLive)
        Else
            'default
            MasterMConLive = New MySqlConnection()
        End If

        Dim ConSQLSlave1 As String = "" & ConfigurationManager.AppSettings("ConSQLSlave1")

        If ConSQLSlave1 <> "" Then
            'webconfig
            MasterMConSlave1 = New MySqlConnection(ConSQLSlave1)
        Else
            'default
            MasterMConSlave1 = New MySqlConnection()
        End If

        Dim ConSQLSlave2 As String = "" & ConfigurationManager.AppSettings("ConSQLSlave2")

        If ConSQLSlave1 <> "" Then
            'webconfig
            MasterMConSlave2 = New MySqlConnection(ConSQLSlave2)
        Else
            'default
            MasterMConSlave2 = New MySqlConnection()
        End If

    End Sub

    Public Function ValidasiUser(ByVal MCon As MySqlConnection, ByVal User As String, ByVal Password As String) As Object()

        Dim Result As Object() = CreateResult()

        Dim Query As String = ""

        Try
            Dim ObjSQL As New ClsSQL

            Dim Process As Boolean = False
            Dim ErrMsg As String = ""

            Query = "Select u.`User`, u.`Password`, u.`Name`, u.`ActiveDate`, u.`InactiveDate`"
            Query &= " , cast(curdate() between u.`ActiveDate` and ifnull(u.`InActiveDate`, curdate()) as char) as IsActive"
            Query &= " , p.`User` is not null as IsPassword, u.`UpdUser`"
            Query &= " From MstUser u"
            Query &= " Left Join MstUser p on (p.`User` = u.`User` And password(p.`Password`) = password(@Password))"
            Query &= " Where u.`User` = @User"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@User", User)
            SqlParam.Add("@Password", Password)

            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)

            If dtQuery Is Nothing Then
                ErrMsg = "Gagal query User"
                GoTo Skip
            End If

            If dtQuery.Rows.Count < 1 Then
                ErrMsg = "User tidak ditemukan"
                GoTo Skip
            End If

            If dtQuery.Rows(0).Item("IsActive").ToString = "1" Then
            Else
                ErrMsg = "Sudah tidak aktif"
                If dtQuery.Rows(0).Item("UpdUser").ToString.Contains("XPASS") Then
                    ErrMsg &= " (XPASS)"
                End If
                GoTo Skip
            End If


            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@User", User)

            If dtQuery.Rows(0).Item("IsPassword").ToString = "1" Then
                'reset counter salah password
                Query = "Delete From MstUserXPass Where `User` = @User"
                ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

            Else

                ErrMsg = "Password salah"

                'counter salah password
                Query = "Insert Into MstUserXPass (`User`, `Qty`, `UpdTime`)"
                Query &= " values ( @User, 1, now() )"
                Query &= " on duplicate key update Qty = Qty + 1, `UpdTime` = now()"
                ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

                'lock user
                Query = "Update MstUser u"
                Query &= " Inner Join MstUserXPass x on (x.`User` = u.`User` and x.Qty >= 5)"
                Query &= " Set u.InActiveDate = date_add(curdate(), interval -1 day), u.UpdTime = now(), u.UpdUser = 'XPASS'"
                Query &= " Where u.`User` = @User"
                ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

                GoTo Skip

            End If 'dari IsPassword

            Process = True

Skip:
            If Process Then
                Result(2) = dtQuery
                Result(1) = ResponseOK
                Result(0) = "0"
            Else
                Result(1) = ErrMsg
            End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog("", "", "ValidasiUser", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

        End Try

        Return Result

    End Function

    Private Function CreateResult() As Object()

        Dim Result As Object()
        ReDim Result(2)

        Result(0) = "9" 'Kode Proses
        Result(1) = "" 'Pesan hasil proses / pesan error
        Result(2) = Nothing 'Variable yang dikembalikan

        Return Result

    End Function

    Public Function ConvertDatatableToString(ByVal dtResult As DataTable) As String

        Return ConvertDatatableToStringBuilder(dtResult, ",", "|", True)

    End Function

    Private Function ConvertDatatableToStringBuilder(ByVal Param As DataTable, ByVal SplitColumn As String, ByVal SplitLine As String, ByVal UseHeader As Boolean) As String

        On Error Resume Next

        Dim sb As New StringBuilder()

        If UseHeader Then

            For i As Integer = 0 To Param.Columns.Count - 1

                sb.Append(Param.Columns(i))

                If i < Param.Columns.Count - 1 Then

                    sb.Append(SplitColumn)

                End If

            Next

            'sb.AppendLine()
            If Param.Rows.Count > 0 Then
                sb.Append(SplitLine)
            End If

        End If

        Dim index As Integer = 0

        Dim MaxIndex As Integer = Param.Rows.Count - 1

        For Each dr As DataRow In Param.Rows

            For i As Integer = 0 To Param.Columns.Count - 1

                sb.Append(dr(i).ToString())

                If i < Param.Columns.Count - 1 Then

                    sb.Append(SplitColumn)

                End If

            Next

            'sb.AppendLine()

            If index < MaxIndex Then

                sb.Append(SplitLine)

            End If

            index = index + 1

        Next

        Return sb.ToString()

    End Function

    Public Function GetServiceTypeList(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString
            Dim ServiceIna As Boolean = False
            Try
                If Param(2).ToString.Trim = "1" Then
                    ServiceIna = True
                End If
            Catch ex As Exception
                ServiceIna = False
            End Try

            Dim ShowSvcCode As Boolean = False
            Try
                If Param(3) = True Then
                    ShowSvcCode = True
                End If
            Catch ex As Exception
                ShowSvcCode = False
            End Try

            MCon = MasterMCon.Clone

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'Else
            Dim ObjSQL As New ClsSQL
            Query = "Select cast(Service as char) as Value"
            If ServiceIna = False Then
                If ShowSvcCode = False Then
                    Query &= " , Cast(Concat(Name, ' (', Alias, ')') As Char) As Display"
                Else
                    Query &= " , Cast(Concat(Name, ' (', Alias, ' / ', Service, ')') As Char) As Display"
                End If
            Else
                If ShowSvcCode = False Then
                    Query &= " , Cast(NameIna As Char) As Display"
                Else
                    Query &= " , Cast(concat(NameIna, ' / ', Service) As Char) As Display"
                End If
            End If
            Query &= " From MstService Where 1=1"
            Query &= " And curdate() between ActiveDate and IfNull(InactiveDate,curdate())"
            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)
            If Not dtQuery Is Nothing Then
                Dim dtTemp As New DataTable
                dtTemp = dtQuery.Clone
                dtTemp.Rows.Add(New String() {"", ""})
                For Each drow As DataRow In dtQuery.Rows
                    dtTemp.ImportRow(drow)
                Next

                Result(2) = ConvertDatatableToString(dtTemp)
                Result(1) = ResponseOK
                Result(0) = "0"
            End If

            'End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetServiceTypeList", User, ex, Query)
            Result(1) &= "Error : " & ex.Message

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

    <WebMethod()>
    Public Function GetECommerceListWithCategoryAndSubCategory(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object(), ByRef dsData As DataSet) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim MCon2 As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            Dim Service As String = ""
            Try
                Service = Param(2).ToString
            Catch ex As Exception
                Service = ""
            End Try

            Dim ServiceAddInfo As String = ""
            Try
                ServiceAddInfo = Param(3).ToString.ToUpper
            Catch ex As Exception
                ServiceAddInfo = ""
            End Try

            Dim OtherCriteria As String = ""
            Try
                OtherCriteria = Param(4).ToString.ToUpper
            Catch ex As Exception
                OtherCriteria = ""
            End Try

            Dim ShowAccountCode As Boolean = False
            Try
                If Param(5) = True Then
                    ShowAccountCode = True
                End If
            Catch ex As Exception
                ShowAccountCode = False
            End Try

            MCon = MasterMCon.Clone
            MCon2 = Master2MCon.Clone

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'Else

            Dim ObjSQL As New ClsSQL

            SqlParam = New Dictionary(Of String, String)

            Query = "(select cast('' as char) as `ValueAccount`, cast('' as char) as `DisplayAccount`  "
            Query &= ", cast('' as char) as `ValueCategory`, cast('' as char) as `DisplayCategory`  "
            Query &= ", cast('' as char) as `ValueSubCategory`, cast('' as char) as `DisplaySubCategory` "
            Query &= ", cast('' as char) as `MappingCategory`, cast('' as char) as `FinFinalProcess`"
            Query &= ") union "
            Query &= "(Select distinct a.`Account` as ValueAccount"
            If ShowAccountCode Then
                Query &= " , cast(concat(a.`Alias`,' (',a.`Account`,')') as char) As DisplayAccount"
            Else
                Query &= " , cast(concat(a.`Alias`,' (',right(a.`Account`,3),')') as char) As DisplayAccount"
            End If
            Query &= " , a.SellerCategory as ValueCategory, c.Description as DisplayCategory"
            Query &= " , a.SellerSubCategory as ValueSubCategory, s.Description as DisplaySubCategory"
            Query &= " , s.Category as MappingCategory"
            Query &= " , a.FinFinalProcess"
            Query &= " From `Account` a"
            Query &= " LEFT JOIN `accountcategory` c on (a.SellerCategory = c.Category)"
            Query &= " LEFT JOIN `accountsubcategory` s on (a.SellerSubCategory = s.SubCategory)"
            If Service <> "" Then
                Query &= " Inner Join `AccountService` v on ("
                Query &= " v.Account = a.Account And v.Service = @Service"
                Query &= " And curdate() between v.ActiveDate and IfNull(v.InactiveDate,curdate())"

                SqlParam.Add("@Service", Service)

                If ServiceAddInfo <> "" Then
                    Query &= " And v.AddInfo like @ServiceAddInfo"
                    SqlParam.Add("@ServiceAddInfo", "%" & ServiceAddInfo & "%")
                End If

                Query &= " )"
            End If

            Select Case OtherCriteria
                Case "REQUESTPICKUP"
                    Query &= " Left Join dashboardaccountmenu          d1 on (d1.Account = a.Account and d1.Menu = 'liRequestPickup')"
                    Query &= " Left Join dashboardaccountusergroupmenu d2 on (d2.Account = a.Account and d2.Menu = 'liRequestPickup')"

                Case "PARTNERITEM"
                    Query &= " inner join partneritem_mstaccount p on ( p.account = a.account"
                    Query &= " And curdate() between p.ActiveDate and IfNull(p.InactiveDate,curdate()) )"

                Case "LAPORANMASALAHPARTNER"
                    Query &= " inner join accountcategory p on ( p.Category = a.SellerCategory"
                    Query &= " And curdate() between p.ActiveDate and IfNull(p.InactiveDate,curdate()) )"
                    Query &= " inner join accountsubcategory q on ( q.SubCategory = a.SellerSubCategory"
                    Query &= " And curdate() between q.ActiveDate and IfNull(q.InactiveDate,curdate()) )"

                Case "DASHBOARD"
                    Query &= " inner join ("
                    Query &= "   select distinct code from mstlogin where type='DSH'"
                    Query &= "   and curdate() between activedate and ifnull(inactivedate, curdate())"
                    Query &= " ) l on (a.account=l.code)"

                Case "INSURANCEREPORT"
                    Query &= " inner join accountinsurancesetting p on (a.account = p.account)"

                Case "REQUESTPICKUPIPPHO"
                    Query &= " inner join ("
                    Query &= "   select distinct code from mstlogin where type='DSH'"
                    Query &= "   and curdate() between activedate and ifnull(inactivedate, curdate())"
                    Query &= " ) l on (a.account=l.code)"
                    Query &= " inner join dashboardaccountmenu d on (a.account = d.account AND d.menu = 'liRequestPickup')"

                Case Else

            End Select

            Query &= " Where a.`Type` in ('1', '2') " 'penanda account type = e-Commerce / Merchant
            Query &= " And curdate() between a.ActiveDate and IfNull(a.InactiveDate,curdate())"

            Select Case OtherCriteria
                Case "REQUESTPICKUP"
                    Query &= " And (d1.Account is not null or d2.Account is not null)"

                Case "LAPORANMASALAHPARTNER"
                    Query &= " And ( "
                    Query &= "    (p.Description like '%procurement%'      or p.Description like '%barang%dagang%')"
                    Query &= " or (q.Description like '%procurement supp%' or q.Description like '%groceries supp%')"
                    Query &= " )"

                Case Else

            End Select

            Query &= " Order By a.`Alias`)"

            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)

            If Not dtQuery Is Nothing Then

                Dim dtCategory As DataTable = dtQuery.DefaultView.ToTable(True, New String() {"ValueCategory"})
                Dim dtSubCategory As DataTable = dtQuery.DefaultView.ToTable(True, New String() {"ValueSubCategory", "DisplaySubCategory", "MappingCategory"})

                If Not dtCategory Is Nothing And Not dtSubCategory Is Nothing Then
                    Dim dtTemp As New DataTable
                    For i As Integer = 0 To dtCategory.Rows.Count - 1
                        dtQuery.DefaultView.RowFilter = "MappingCategory IN ('" & dtCategory.Rows(i).Item("ValueCategory") & "')"
                        dtTemp = dtQuery.DefaultView.ToTable
                        dtTemp.TableName = "CAT_" & IIf(dtCategory.Rows(i).Item("ValueCategory") = "", "0", dtCategory.Rows(i).Item("ValueCategory"))
                        dsData.Tables.Add(dtTemp)
                    Next

                    For i As Integer = 0 To dtCategory.Rows.Count - 1
                        dtSubCategory.DefaultView.RowFilter = "MappingCategory IN ('" & dtCategory.Rows(i).Item("ValueCategory") & "')"
                        dtTemp = dtSubCategory.DefaultView.ToTable
                        dtTemp.TableName = "MAPPING_" & IIf(dtCategory.Rows(i).Item("ValueCategory") = "", "0", dtCategory.Rows(i).Item("ValueCategory"))
                        dsData.Tables.Add(dtTemp)
                    Next

                    For i As Integer = 0 To dtSubCategory.Rows.Count - 1
                        dtQuery.DefaultView.RowFilter = "ValueSubCategory IN ('" & dtSubCategory.Rows(i).Item("ValueSubCategory") & "')"
                        dtTemp = dtQuery.DefaultView.ToTable
                        dtTemp.TableName = "SUBCAT_" & IIf(dtSubCategory.Rows(i).Item("ValueSubCategory") = "", "0", dtSubCategory.Rows(i).Item("ValueSubCategory"))
                        dsData.Tables.Add(dtTemp)
                    Next
                End If

                Result(2) = ConvertDatatableToString(dtQuery)
                Result(1) = ResponseOK
                Result(0) = "0"

            End If

            'End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetECommerceListWithCategoryAndSubCategory", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

        Finally
            If MCon.State <> ConnectionState.Closed Then
                MCon.Close()
            End If
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            Try
                MCon2.Close()
            Catch ex As Exception
            End Try
            Try
                MCon2.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result

    End Function

    <WebMethod()>
    Public Function GetECommerceList(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim MCon2 As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            Dim Service As String = ""
            Try
                Service = Param(2).ToString
            Catch ex As Exception
                Service = ""
            End Try

            Dim ServiceAddInfo As String = ""
            Try
                ServiceAddInfo = Param(3).ToString.ToUpper
            Catch ex As Exception
                ServiceAddInfo = ""
            End Try

            Dim OtherCriteria As String = ""
            Try
                OtherCriteria = Param(4).ToString.ToUpper
            Catch ex As Exception
                OtherCriteria = ""
            End Try

            Dim ShowAccountCode As Boolean = False
            Try
                If Param(5) = True Then
                    ShowAccountCode = True
                End If
            Catch ex As Exception
                ShowAccountCode = False
            End Try

            MCon = MasterMCon.Clone
            MCon2 = Master2MCon.Clone

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'Else

            Dim ObjSQL As New ClsSQL

            SqlParam = New Dictionary(Of String, String)

            Query = "Select distinct a.`Account` as Value"
            If ShowAccountCode = False Then
                Query &= " , cast(concat(a.`Alias`,' (',right(a.`Account`,3),')') as char) As Display"
            Else
                Query &= " , cast(concat(a.`Alias`,' (', a.`Account` ,')') as char) As Display"
            End If
            Query &= " From `Account` a"
            If Service <> "" Then
                Query &= " Inner Join `AccountService` v on ("
                Query &= " v.Account = a.Account And v.Service = @Service"
                Query &= " And curdate() between v.ActiveDate and IfNull(v.InactiveDate,curdate())"

                SqlParam.Add("@Service", Service)

                If ServiceAddInfo <> "" Then
                    Query &= " And v.AddInfo like @ServiceAddInfo"
                    SqlParam.Add("@ServiceAddInfo", "%" & ServiceAddInfo & "%")
                End If
                Query &= " )"
            End If

            Select Case OtherCriteria
                Case "REQUESTPICKUP"
                    Query &= " Left Join dashboardaccountmenu          d1 on (d1.Account = a.Account and d1.Menu = 'liRequestPickup')"
                    Query &= " Left Join dashboardaccountusergroupmenu d2 on (d2.Account = a.Account and d2.Menu = 'liRequestPickup')"

                Case "PARTNERITEM"
                    Query &= " inner join partneritem_mstaccount p on ( p.account = a.account"
                    Query &= " And curdate() between p.ActiveDate and IfNull(p.InactiveDate,curdate()) )"

                Case "LAPORANMASALAHPARTNER"
                    Query &= " inner join accountcategory p on ( p.Category = a.SellerCategory"
                    Query &= " And curdate() between p.ActiveDate and IfNull(p.InactiveDate,curdate()) )"
                    Query &= " inner join accountsubcategory q on ( q.SubCategory = a.SellerSubCategory"
                    Query &= " And curdate() between q.ActiveDate and IfNull(q.InactiveDate,curdate()) )"
                Case "DASHBOARD"
                    Query &= " inner join ("
                    Query &= "   select distinct code from mstlogin where type='DSH'"
                    Query &= "   and curdate() between activedate and ifnull(inactivedate, curdate())"
                    Query &= " ) l on (a.account=l.code)"

                Case "INSURANCEREPORT"
                    Query &= " inner join accountinsurancesetting p on (a.account = p.account)"

                Case "REQUESTPICKUPIPPHO"
                    Query &= " inner join ("
                    Query &= "   select distinct code from mstlogin where type='DSH'"
                    Query &= "   and curdate() between activedate and ifnull(inactivedate, curdate())"
                    Query &= " ) l on (a.account=l.code)"
                    Query &= " inner join dashboardaccountmenu d on (a.account = d.account AND d.menu = 'liRequestPickup')"
                Case Else
            End Select
            Query &= " Where a.`Type` in ('1', '2') " 'penanda account type = e-Commerce / Merchant
            Query &= " And curdate() between a.ActiveDate and IfNull(a.InactiveDate,curdate())"
            Select Case OtherCriteria
                Case "REQUESTPICKUP"
                    Query &= " And (d1.Account is not null or d2.Account is not null)"
                Case "LAPORANMASALAHPARTNER"
                    Query &= " And ( "
                    Query &= "    (p.Description like '%procurement%'      or p.Description like '%barang%dagang%')"
                    Query &= " or (q.Description like '%procurement supp%' or q.Description like '%groceries supp%')"
                    Query &= " )"

                Case Else

            End Select
            Query &= " Order By a.`Alias`"
            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon2, Query, SqlParam)
            If Not dtQuery Is Nothing Then
                Dim dtTemp As New DataTable
                dtTemp = dtQuery.Clone
                dtTemp.Rows.Add(New String() {"", ""})
                For Each drow As DataRow In dtQuery.Rows
                    dtTemp.ImportRow(drow)
                Next
                Result(2) = ConvertDatatableToString(dtTemp)
                Result(1) = ResponseOK
                Result(0) = "0"

            End If

            'End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetECommerceList", User, ex, Query)
            Result(1) &= "Error : " & ex.Message
        Finally
            If MCon.State <> ConnectionState.Closed Then
                MCon.Close()
            End If
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            Try
                MCon2.Close()
            Catch ex As Exception
            End Try
            Try
                MCon2.Dispose()
            Catch ex As Exception
            End Try
        End Try
        Return Result
    End Function
    <WebMethod()>
    Public Function GetECommerceListByFinType(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()
        Dim MCon As New MySqlConnection
        Dim Query As String = ""
        Dim User As String = Param(0).ToString
        Try
            Dim Password As String = Param(1).ToString
            Dim FinType As String = "7"
            Try
                FinType = Param(2).ToString '7 = Weekly, 30 = Monthly
            Catch ex As Exception
                FinType = "7"
            End Try


            MCon = MasterMCon.Clone

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            If UserOK(0) <> "0" Then
                Result(1) = UserOK(1)
            Else

                Dim ObjSQL As New ClsSQL

                Query = "Select `Account` as Value"
                Query &= " , cast(concat(`Alias`,' (',right(`Account`,3),')') as char(50)) As Display"
                Query &= " From `Account` Where 1=1"
                Query &= " And curdate() between ActiveDate and IfNull(InactiveDate,curdate())"
                Query &= " And `Type` in ('1', '2') " 'penanda account type = e-Commerce / Merchant
                Query &= " And FinFinalProcess = '" & FinType & "'"
                Query &= " Order By `Alias`"

                Dim dtQuery As DataTable = ObjSQL.ExecDatatable(MCon, Query)

                If Not dtQuery Is Nothing Then

                    Dim dtTemp As New DataTable

                    dtTemp = dtQuery.Clone

                    dtTemp.Rows.Add(New String() {"", ""})

                    For Each drow As DataRow In dtQuery.Rows
                        dtTemp.ImportRow(drow)
                    Next

                    Result(2) = ConvertDatatableToString(dtTemp)
                    Result(1) = ResponseOK
                    Result(0) = "0"

                End If

            End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetECommerceListByFinType", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Public Function GetHubList(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim MCon2 As New MySqlConnection

        Dim Query As String = ""

        Dim User As String = Param(0).ToString
        Try
            Dim Password As String = Param(1).ToString

            Dim OrderBy As String = ""
            Try
                OrderBy = Param(2).ToString.Trim.ToUpper
            Catch ex As Exception
                OrderBy = ""
            End Try

            Dim AddInfo As String = ""
            Try
                AddInfo = Param(3).ToString.Trim.ToUpper
            Catch ex As Exception
                AddInfo = ""
            End Try

            MCon = MasterMCon.Clone
            MCon2 = Master2MCon.Clone

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'Else

            Dim ObjSQL As New ClsSQL

            Query = "Select * From ("
            Query &= " Select Hub as Value"
            Query &= " , Cast(Concat(Name, ' (', Alias, ' - ', Hub, ')') As Char(50)) As Display"
            Query &= " , Alias as HubAlias, Region as RegionIPP"
            Query &= " From MstHub Where 1=1"
            Query &= " And curdate() between ActiveDate and IfNull(InactiveDate,curdate())"
            If OrderBy = "REGIONIPP" Then
                Query &= " Order By Region, Name"
            Else
                Query &= " Order By Name"
            End If
            Query &= " ) x"

            If AddInfo.Contains("ADDITIONALHUB=Y") Then
                Query &= " Union"
                Query &= " Select * From ("
                Query &= " Select Hub as Value"
                Query &= " , Cast(Concat(Name, ' (', Alias, ' - ', Hub, ')') As Char(50)) As Display"
                Query &= " , Alias as HubAlias, Region as RegionIPP"
                Query &= " From MstHub_Additional"
                Query &= " ) y"
            End If

            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon2, Query, Nothing)

            If Not dtQuery Is Nothing Then

                Dim dtTemp As New DataTable

                dtTemp = dtQuery.Clone
                dtTemp.Rows.Add(New String() {"", ""})
                For Each drow As DataRow In dtQuery.Rows
                    dtTemp.ImportRow(drow)
                Next
                Result(2) = ConvertDatatableToString(dtTemp)
                Result(1) = ResponseOK
                Result(0) = "0"
            End If

            'End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetHubList", User, ex, Query)
            Result(1) &= "Error : " & ex.Message
        Finally
            If MCon.State <> ConnectionState.Closed Then
                MCon.Close()
            End If
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            Try
                MCon2.Close()
            Catch ex As Exception
            End Try
            Try
                MCon2.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result

    End Function

    Public Function GetTrackingStatusListForReport(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString
            MCon = MasterMCon.Clone

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            If UserOK(0) <> "0" Then
                Result(1) = UserOK(1)
            Else

                Dim ObjSQL As New ClsSQL

                Query = "Select distinct `Status` as Value"
                Query &= " , cast(concat(`Description`,' (',`Status`,')') as char(50)) As Display"
                Query &= " From `MstTrackingStatus`"
                Query &= " Where FlgReport = '1'"

                Dim dtQuery As DataTable = ObjSQL.ExecDatatable(MCon, Query)

                If Not dtQuery Is Nothing Then

                    Dim dtTemp As New DataTable

                    dtTemp = dtQuery.Clone

                    dtTemp.Rows.Add(New String() {"", ""})

                    For Each drow As DataRow In dtQuery.Rows
                        dtTemp.ImportRow(drow)
                    Next

                    Result(2) = ConvertDatatableToString(dtTemp)
                    Result(1) = ResponseOK
                    Result(0) = "0"

                End If

            End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetTrackingStatusList", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    <WebMethod()>
    Public Function GetRegionList(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim MCon2 As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString
        Try
            Dim Password As String = Param(1).ToString

            MCon = MasterMCon.Clone
            MCon2 = Master2MCon.Clone

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'Else

            Dim ObjSQL As New ClsSQL

            Query = "Select Region as Value"
            Query &= " , Cast(Concat(Name, ' - ', replace(case when alias = '' then 'N/A' else alias END, ',' , ' '), ' (', Region , ')') As Char(50)) As Display"
            Query &= " From MstRegion"
            Query &= " Where curdate() between ActiveDate and IfNull(InactiveDate,curdate())"
            Query &= " Order By CAST(Region AS UNSIGNED), Region"

            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)

            If Not dtQuery Is Nothing Then

                Dim dtTemp As New DataTable

                dtTemp = dtQuery.Clone

                dtTemp.Rows.Add(New String() {"", ""})

                For Each drow As DataRow In dtQuery.Rows
                    dtTemp.ImportRow(drow)
                Next

                Result(2) = ConvertDatatableToString(dtTemp)
                Result(1) = ResponseOK
                Result(0) = "0"

            End If

            'End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetRegionList", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

        Finally
            If MCon.State <> ConnectionState.Closed Then
                MCon.Close()
            End If
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            Try
                MCon2.Close()
            Catch ex As Exception
            End Try
            Try
                MCon2.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result

    End Function
    <WebMethod()>
    Public Function LoginPasswordUpdateV2(ByVal Param As Object()) As Object()
        Dim Result As Object() = CreateResult()
        Dim MCon As New MySqlConnection
        Dim Query As String = ""
        Dim User As String = Param(0).ToString
        Try
            Dim Password As String = Param(1).ToString

            Dim LoginUser As String = Param(2).ToString
            Dim LoginType As String = Param(3).ToString
            Dim LoginPassword As String = Param(4).ToString
            Dim LoginNewPassword As String = Param(5).ToString
            MCon = MasterMCon.Clone
            MCon.Open()
            Dim ObjSQL As New ClsSQL

            Dim Process As Boolean = False
            Dim ErrMsg As String = ""

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            If UserOK(0) <> "0" Then
                ErrMsg = UserOK(1)
                GoTo Skip
            End If

            If LoginUser.Trim = "" Or LoginType.Trim = "" Or LoginPassword.Trim = "" Or LoginNewPassword.Trim = "" Then
                ErrMsg = "Input User/Type/Password/Password baru"
                GoTo Skip
            End If

            If UCase(LoginPassword.Trim) = UCase(LoginNewPassword.Trim) Then
                ErrMsg = "Password baru tidak boleh sama dengan Password saat ini"
                GoTo Skip
            End If

            Dim mLength As Short = 6
            If LoginNewPassword.Trim.Length < mLength Then
                ErrMsg = "Password minimal " & mLength & " karakter"
                GoTo Skip
            End If

            Select Case LoginType
                Case "DSH"

                Case "RPT"
                    Dim isAlphaAndNumeric As Boolean = System.Text.RegularExpressions.Regex.IsMatch(LoginNewPassword.Trim, "^(?=.*[0-9])(?=.*[a-zA-Z])([a-zA-Z0-9]+)$")
                    If Not isAlphaAndNumeric Then
                        ErrMsg = "Password harus kombinasi huruf dan angka"
                        GoTo Skip
                    End If

                Case "HUB"
                    Dim isAlphaAndNumeric As Boolean = System.Text.RegularExpressions.Regex.IsMatch(LoginNewPassword.Trim, "^(?=.*[0-9])(?=.*[a-zA-Z])([a-zA-Z0-9]+)$")
                    If Not isAlphaAndNumeric Then
                        ErrMsg = "Password harus kombinasi huruf dan angka"
                        GoTo Skip
                    End If

                Case "HCR"
                    Dim isAlphaAndNumeric As Boolean = System.Text.RegularExpressions.Regex.IsMatch(LoginNewPassword.Trim, "^(?=.*[0-9])(?=.*[a-zA-Z])([a-zA-Z0-9]+)$")
                    If Not isAlphaAndNumeric Then
                        ErrMsg = "Password harus kombinasi huruf dan angka"
                        GoTo Skip
                    End If

                Case Else

            End Select

            Query = "Select `User` From MstLogin"
            Query &= " Where `User` = @LoginUser"
            Query &= " And `Type` = @LoginType"
            Query &= " And password(`Password`) = password(@LoginPassword)"
            Query &= " And curdate() between ActiveDate and IfNull(InactiveDate,curdate())"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@LoginUser", LoginUser)
            SqlParam.Add("@LoginType", LoginType)
            SqlParam.Add("@LoginPassword", LoginPassword)

            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)

            If dtQuery Is Nothing Then
                ErrMsg = "Gagal query"
                GoTo Skip
            End If

            If dtQuery.Rows.Count < 1 Then
                ErrMsg = "User tidak ditemukan / Password salah / sudah Tidak Aktif"
                GoTo Skip
            End If

            Query = "Update MstLogin"
            Query &= " Set `Password` = @LoginNewPassword"
            Query &= " , ResetPIN = 0"
            Query &= " , UpdTime = now(), UpdUser = 'PassUpd'"
            Query &= " Where User = @LoginUser"
            Query &= " And `Type` = @LoginType"
            Query &= " And password(`Password`) = password(@LoginPassword)"
            Query &= " And curdate() between ActiveDate and IfNull(InactiveDate,curdate())"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@LoginNewPassword", LoginNewPassword)
            SqlParam.Add("@LoginUser", LoginUser)
            SqlParam.Add("@LoginType", LoginType)
            SqlParam.Add("@LoginPassword", LoginPassword)

            If ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam) = False Then
                ErrMsg = "Update password gagal"
                GoTo Skip
            End If

            Query = " Delete From MstLoginXPass Where User = @LoginUser And `Type` = @LoginType"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@LoginUser", LoginUser)
            SqlParam.Add("@LoginType", LoginType)

            ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

            Process = True

Skip:

            If Process Then
                Result(2) = ""
                Result(1) = ResponseOK
                Result(0) = "0"
            Else
                Result(1) = ErrMsg
            End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog("", "", "LoginPasswordUpdateV2", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    <WebMethod()>
    Public Function QueryTools(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""
        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString
            Query = Param(2).ToString.Trim

            If Param(3) = "dblive" Then
                MCon = MasterMCon.Clone
            ElseIf Param(3) = "dbslave1" Then
                MCon = MasterMConSlave1.Clone
            Else
                MCon = MasterMConSlave2.Clone
            End If

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'GoTo Skip
            'End If

            If Query = "" Then
                Result(1) = "Input Query"
                GoTo Skip
            End If

            Dim ObjSQL As New ClsSQL
            Dim dtQuery As DataTable = Nothing

            Dim IsValid As Boolean = True
            Dim Response1 As String = ""

            If Query.ToUpper.StartsWith("SELECT") _
            Or Query.ToUpper.StartsWith("SHOW TABLES") _
            Or Query.ToUpper.StartsWith("CALL") _
            Or Query.ToUpper.StartsWith("SHOW SLAVE STATUS") Then

                dtQuery = ObjSQL.ExecDatatable(MCon, Query) 'balikin data
                If dtQuery Is Nothing Then
                    Result(1) = "Gagal query"
                    GoTo Skip
                End If

            ElseIf Query.ToUpper.StartsWith("SHOW CREATE TABLE") Then
                Dim dtTemp As DataTable = ObjSQL.ExecDatatable(MCon, Query)
                If dtTemp Is Nothing Then
                    Result(1) = "Gagal query"
                    GoTo Skip
                End If

                Response1 = dtTemp.Rows(0).Item("Create Table").ToString 'balikin query create table

            ElseIf Query.ToUpper.StartsWith("SHOW PROCEDURE STATUS") Then

                Dim dtTemp As DataTable = ObjSQL.ExecDatatable(MCon, Query)
                If dtTemp Is Nothing Then
                    Result(1) = "Gagal query"
                    GoTo Skip
                End If

                dtQuery = New DataTable
                dtQuery.Columns.Add("Db")
                dtQuery.Columns.Add("Name")

                For t As Integer = 0 To dtTemp.Rows.Count - 1
                    dtQuery.Rows.Add()
                    dtQuery.Rows(t).Item("Db") = dtTemp.Rows(t).Item("Db").ToString
                    dtQuery.Rows(t).Item("Name") = dtTemp.Rows(t).Item("Name").ToString
                Next
                dtQuery.AcceptChanges()

            ElseIf Query.ToUpper.StartsWith("SHOW CREATE PROCEDURE") Then
                Dim dtTemp As DataTable = ObjSQL.ExecDatatable(MCon, Query)
                If dtTemp Is Nothing Then
                    Result(1) = "Gagal query"
                    GoTo Skip
                End If

                Response1 = dtTemp.Rows(0).Item("Create Procedure").ToString 'balikin query create procedure

            ElseIf Query.ToUpper.StartsWith("SHOW SLAVE STATUS") Then
                dtQuery = ObjSQL.ExecDatatable(MCon, Query)
                If dtQuery Is Nothing Then
                    Result(1) = "Gagal query"
                    GoTo Skip
                End If

            ElseIf Query.ToUpper.StartsWith("INSERT") _
            Or Query.ToUpper.StartsWith("UPDATE") _
            Or Query.ToUpper.StartsWith("DELETE") Then

                Dim MCom As New MySqlCommand("", MCon)
                MCom.CommandText = Query

                Try
                    Response1 = MCom.ExecuteNonQuery
                Catch ex As Exception
                    Response1 = ex.Message
                End Try

            Else
                IsValid = False
            End If

            If IsValid = False Then
                Result(1) = "Query tidak boleh dijalankan"
                GoTo Skip
            End If


            If Not dtQuery Is Nothing Then
                Result(2) = ConvertDatatableToStringWithBarcode(dtQuery)
            End If
            Result(1) = IIf(Response1 <> "", Response1, ResponseOK)
            Result(0) = "0"

Skip:

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "QueryTools", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Public Function ConvertDatatableToStringWithBarcode(ByVal dtResult As DataTable) As String

        Return ConvertDatatableToStringBuilder(dtResult, "#,#", "#|#", True)

        On Error Resume Next

        'Dim Result As String = ""

        ''column header
        'For j As Integer = 0 To dtResult.Columns.Count - 1
        '    If j <> 0 Then
        '        Result &= "#,#"
        '    End If
        '    Result &= dtResult.Columns(j).ColumnName
        'Next

        'For i As Integer = 0 To dtResult.Rows.Count - 1

        '    If Result <> "" Then 'pemisah antar row
        '        Result &= "#|#"
        '    End If

        '    For j As Integer = 0 To dtResult.Columns.Count - 1
        '        If j <> 0 Then
        '            Result &= "#,#"
        '        End If
        '        Result &= dtResult.Rows(i).Item(j).ToString
        '    Next
        'Next

        'Return Result

    End Function

    Public Function UpdateLastMileExpeditionList(ByVal AppName As String, ByVal AppVersion As String _
    , ByVal Param As Object()) As Object()

        'untuk download data terakhir, pakai sp report_LastMileExpedition

        Dim Method As String = "UpdateLastMileExpeditionList"
        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            Dim LastMileExpeditionList As String = Param(2).ToString.Trim.ToUpper 'ProvinceCode,ExpeditionCode,DstHub,Service|...
            Dim UserID As String = Param(3).ToString.Trim.ToUpper

            Dim MTrn As MySqlTransaction

            MCon = MasterMCon.Clone

            Dim MCom As New MySqlCommand("", MCon)
            MCon.Open()

            MTrn = MCon.BeginTransaction()
            MCom.Transaction = MTrn

            Dim Process As Boolean = False

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'GoTo Skip
            'End If

            If LastMileExpeditionList = "" Then
                Result(1) = "Input daftar Ekspedisi LastMile"
                GoTo Skip
            End If

            Dim ReqParam() As String
            ReDim ReqParam(Param.Length - 1)
            ReqParam(0) = Param(0)
            ReqParam(1) = Param(1)
            ReqParam(2) = "" 'Param(2)
            ReqParam(3) = Param(3)

            CreateRequestLog(AppName, AppVersion, Method, User, "", ReqParam)

            Dim LastMileExpeditionListSplit As String() = LastMileExpeditionList.Split("|")

            'Susun daftar
            Query = ""
            For i As Integer = 0 To LastMileExpeditionListSplit.Length - 1
                If LastMileExpeditionListSplit(i) <> "" Then
                    Dim Province As String = LastMileExpeditionListSplit(i).Split(",")(0).Trim
                    Dim Expedition As String = LastMileExpeditionListSplit(i).Split(",")(1).Trim
                    Dim DstHub As String = ""
                    Dim Priority As Integer = LastMileExpeditionListSplit(i).Split(",")(4).Trim
                    Try
                        DstHub = LastMileExpeditionListSplit(i).Split(",")(2).Trim.ToUpper
                    Catch ex As Exception
                        DstHub = ""
                    End Try
                    Dim Service As String = ""
                    Try
                        Service = LastMileExpeditionListSplit(i).Split(",")(3).Trim
                    Catch ex As Exception
                        Service = ""
                    End Try

                    If Query <> "" Then
                        Query &= ","
                    End If
                    Query &= "( '" & Province & "', '" & Expedition & "', '" & DstHub & "', '" & Service & "', '" & Priority & "' )"
                End If
            Next


            Dim TableName As String = "MstLastMileExpedition2"

            'Temporary Table
            Dim TempTable As String = "Temp" & TableName

            MCom.CommandText = "Drop Temporary Table If Exists " & TempTable
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            MCom.CommandText = "Create Temporary Table " & TempTable & " Like " & TableName
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            MCom.CommandText = "Alter Table " & TempTable
            MCom.CommandText &= " Drop Primary Key"
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            MCom.CommandText = "Alter Table " & TempTable
            MCom.CommandText &= " Add Index i_idx (`Province`,`Expedition`,`DstHub`,`Service`)"
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            MCom.CommandText = "Insert Into " & TempTable & " ("
            MCom.CommandText &= " Province, Expedition, DstHub, Service, Priority"
            MCom.CommandText &= " ) values"
            MCom.CommandText &= " " & Query
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            MCom.CommandText = "Delete From " & TempTable
            MCom.CommandText &= " Where Province = '' or Expedition = '' or DstHub = '' or Service = ''"
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            '#Check duplicate data
            MCom.CommandText = "Select ifnull(group_concat(DataId),'') From ("
            MCom.CommandText &= "   Select concat(Province,'.',DstHub,'.',Service,'.',Expedition) as DataId, count(*)"
            MCom.CommandText &= "   From " & TempTable
            MCom.CommandText &= "   Group By concat(Province,'.',DstHub,'.',Service,'.',Expedition)"
            MCom.CommandText &= "   Having count(*) > 1"
            MCom.CommandText &= "   Limit 10"
            MCom.CommandText &= " )x"
            Dim Validasi As String = ("" & MCom.ExecuteScalar).ToString.Trim
            If Validasi <> "" Then
                Result(1) = "Ada data lebih dari 1 baris (" & Validasi & ")"
                GoTo Skip
            End If

            '#Check max 2 baris setting
            MCom.CommandText = "Select ifnull(group_concat(DataId),'') From ("
            MCom.CommandText &= "   Select concat(Province,'.',DstHub,'.',Service) as DataId, count(*)"
            MCom.CommandText &= "   From " & TempTable
            MCom.CommandText &= "   Group By concat(Province,'.',DstHub,'.',Service)"
            MCom.CommandText &= "   Having count(*) > 2"
            MCom.CommandText &= "   Limit 10"
            MCom.CommandText &= " )x"
            Dim ValidasiCount As String = ("" & MCom.ExecuteScalar).ToString.Trim
            If Validasi <> "" Then
                Result(1) = "Ada data lebih dari 2 baris (" & ValidasiCount & ")"
                GoTo Skip
            End If

            MCom.CommandText = "Select ifnull(group_concat(Province),'') From ("
            MCom.CommandText &= "   Select Province"
            MCom.CommandText &= "   From " & TempTable
            MCom.CommandText &= "   Where Province not in (Select Province From mstprovince)"
            MCom.CommandText &= "   Limit 10"
            MCom.CommandText &= " )x"
            Validasi = ("" & MCom.ExecuteScalar).ToString.Trim
            If Validasi <> "" Then
                Result(1) = "Propinsi tidak terdaftar (" & Validasi & ")"
                GoTo Skip
            End If

            MCom.CommandText = "Select ifnull(group_concat(DstHub),'') From ("
            MCom.CommandText &= "   Select DstHub"
            MCom.CommandText &= "   From " & TempTable
            MCom.CommandText &= "   Where DstHub not in (Select Hub From MstHub UNION Select '*' as Hub)"
            MCom.CommandText &= "   Limit 10"
            MCom.CommandText &= " )x"
            Validasi = ("" & MCom.ExecuteScalar).ToString.Trim
            If Validasi <> "" Then
                Result(1) = "Hub tidak terdaftar (" & Validasi & ")"
                GoTo Skip
            End If

            MCom.CommandText = "Select ifnull(group_concat(Service),'') From ("
            MCom.CommandText &= "   Select Service"
            MCom.CommandText &= "   From " & TempTable
            MCom.CommandText &= "   Where Service not in (Select Service From MstService UNION Select '*' as Service)"
            MCom.CommandText &= "   Limit 10"
            MCom.CommandText &= " )x"
            Validasi = ("" & MCom.ExecuteScalar).ToString.Trim
            If Validasi <> "" Then
                Result(1) = "Service tidak terdaftar (" & Validasi & ")"
                GoTo Skip
            End If

            MCom.CommandText = "Select ifnull(group_concat(Expedition),'') From ("
            MCom.CommandText &= "   Select Expedition"
            MCom.CommandText &= "   From " & TempTable
            MCom.CommandText &= "   Where Expedition not in (Select Account From account Where `Type`='3')"
            MCom.CommandText &= "   Limit 10"
            MCom.CommandText &= " )x"
            Validasi = ("" & MCom.ExecuteScalar).ToString.Trim
            If Validasi <> "" Then
                Result(1) = "Ekspedisi tidak terdaftar (" & Validasi & ")"
                GoTo Skip
            End If

            MCom.CommandText = "Select ifnull(group_concat(Priority),'') From ("
            MCom.CommandText &= " Select Priority"
            MCom.CommandText &= " From " & TempTable
            MCom.CommandText &= " Where Priority NOT IN (1,2)"
            MCom.CommandText &= " Limit 10"
            MCom.CommandText &= ")x"
            Validasi = ("" & MCom.ExecuteScalar).ToString.Trim
            If Validasi <> "" Then
                Result(1) = "Input Priority tidak sesuai (" & Validasi & ")"
                GoTo Skip
            End If

            'New Table
            Dim NewTable As String = TableName & "_New"

            MCom.CommandText = "Drop Table If Exists " & NewTable
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            MCom.CommandText = "Create Table " & NewTable & " Like " & TableName
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            MCom.CommandText = "Insert Into " & NewTable
            MCom.CommandText &= " Select * From " & TempTable
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            MCom.CommandText = "Update " & NewTable
            MCom.CommandText &= " Set UpdTime = now(), UpdUser = '" & UserID & "'"
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            'BackupTable
            Dim BackupTable As String = TableName & "_Backup"
            MCom.CommandText = "Drop table if exists " & BackupTable
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try


            MCom.CommandText = "Rename Table " & TableName & " to " & BackupTable
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            MCom.CommandText = "Rename Table " & NewTable & " to " & TableName
            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, MCom.CommandText)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            Process = True

Skip:

            If Process Then
                MTrn.Commit()

                Result(1) = ResponseOK
                Result(0) = "0"

                Try
                    Result(2) = ""

                    Query = "Select Province, DstHub, Service, Expedition, Count(*) as nData"
                    Query &= " From " & TempTable
                    Query &= " Group By Province, DstHub, Service, Expedition"

                    Dim ObjSQL As New ClsSQL
                    Dim dtData As DataTable = ObjSQL.ExecDatatable(MCon, Query)

                    Dim nData As Integer = 0

                    For i As Integer = 0 To dtData.Rows.Count - 1
                        If Result(2) <> "" Then
                            Result(2) &= vbCrLf
                        End If
                        Result(2) &= "> Propinsi " & dtData.Rows(i).Item("Province").ToString
                        Result(2) &= ", DstHub " & dtData.Rows(i).Item("DstHub").ToString
                        Result(2) &= ", Service " & dtData.Rows(i).Item("Service").ToString
                        Result(2) &= ", Expedition " & dtData.Rows(i).Item("Expedition").ToString
                        Result(2) &= " : " & dtData.Rows(i).Item("nData").ToString & " baris"

                        nData += CInt(dtData.Rows(i).Item("nData"))
                    Next

                    Result(2) = "Hasil Proses" & vbCrLf & Result(2)

                Catch ex As Exception
                    Result(2) = ""
                End Try

            Else
                MTrn.Rollback()
            End If

            CreateResponseLog(AppName, AppVersion, Method, User, Result)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Public Sub CreateResponseLog(ByVal AppName As String, ByVal AppVersion As String _
    , ByVal Method As String, ByVal User As String, ByVal Result As Object(), Optional ByVal ReferenceNo As String = "" _
    , Optional ByVal Keyword As String = "")

        On Error Resume Next

        Dim RspLog As String = ""

        For i As Integer = 0 To Result.Length - 1
            If i > 0 Then
                RspLog &= "|"
            End If
            'If RspLog <> "" Then
            '    RspLog &= "|"
            'End If
            RspLog &= StrSQLRemoveChar(Result(i).ToString)
        Next

        Dim e As New ClsError
        e.ResponseLog(MasterMCon, AppVersion, Method, User, StrSQLRemoveChar(ReferenceNo) & " " & RspLog, Keyword)

    End Sub
    Public Sub CreateRequestLog(ByVal AppName As String, ByVal AppVersion As String _
    , ByVal Method As String, ByVal User As String, ByVal Code As String, ByVal Param As Object() _
    , Optional ByVal Keyword As String = "")

        On Error Resume Next

        Dim ReqLog As String = ""

        '0 dan 1 = User dan Password
        For i As Integer = 2 To Param.Length - 1
            If i > 2 Then
                ReqLog &= "|"
            End If
            'If ReqLog <> "" Then
            '    ReqLog &= "|"
            'End If
            ReqLog &= StrSQLRemoveChar(Param(i).ToString)
        Next

        Dim e As New ClsError
        e.RequestLog(MasterMCon, AppName, AppVersion, Method, User, Code & " " & ReqLog, Keyword)

    End Sub

    Public Function StrSQLRemoveChar(ByVal Ori As String) As String

        Dim Result As String = Ori

        Result = Result.Replace("'", "").Replace(";", "").Replace("\", " ").Replace("""", " ").Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ").Replace("’", "")

        'remove ASCII extended character 
        For i As Integer = 128 To 255
            Result = Result.Replace(Chr(i), " ")
        Next

        While Result.Contains("  ")
            Result = Result.Replace("  ", " ")
        End While

        Result = Result.Trim

        Return Result

    End Function

    Public Function LiveTrackingUpdate(ByVal MCon As MySqlConnection, ByVal Param() As Object) As Object()

        Dim Method As String = "LiveTrackingUpdate"

        Dim Result As Object() = CreateResult()

        Dim Query As String = ""

        Dim MCom As New MySqlCommand("", MCon) 'antisipasi ada begin tran pada caller

        Try
            Dim TrackNum As String = Param(2).ToString.Trim.ToUpper

            CreateRequestLog(wsAppName, wsAppVersion, Method, "SYS", "", Param, TrackNum)

            Dim ThirdParty As String = Param(3).ToString.Trim

            Dim DriverID As String = ""
            Try
                DriverID = Param(4).ToString.Trim.ToUpper
            Catch ex As Exception
                DriverID = ""
            End Try

            Dim DriverName As String = ""
            Try
                DriverName = Param(5).ToString.Trim.ToUpper
            Catch ex As Exception
                DriverName = ""
            End Try

            Dim DriverPhone As String = ""
            Try
                DriverPhone = Param(6).ToString.Trim
            Catch ex As Exception
                DriverPhone = ""
            End Try

            Dim VehicleNo As String = ""
            Try
                VehicleNo = Param(7).ToString.Trim.ToUpper
            Catch ex As Exception
                VehicleNo = ""
            End Try

            Dim TrackingUrl As String = ""
            Try
                TrackingUrl = Param(8).ToString.Trim
            Catch ex As Exception
                TrackingUrl = ""
            End Try


            Dim sb As New StringBuilder
            sb.Append("Insert Into LiveTracking (")
            sb.Append(" TrackNum, OrderNo, Account")
            sb.Append(" , DriverID, DriverName, DriverPhone")
            sb.Append(" , VehicleNo, TrackingUrl")
            sb.Append(" , PushStatus, RetryPush")
            sb.Append(" , TrackTime, TrackUser")
            sb.Append(" ) Select")
            sb.Append(" t.TrackNum, t.OrderNo, (case when t.PushAccount <> '' then t.PushAccount else t.ShAccount end)")
            sb.Append(" , @DriverID, @DriverName, @DriverPhone")
            sb.Append(" , @VehicleNo, @TrackingUrl")
            sb.Append(" , '0', '0'")
            sb.Append(" , now(), @ThirdParty")
            sb.Append(" From `Transaction` t Where t.TrackNum = @TrackNum")
            Query = sb.ToString

            MCom.CommandText = Query

            MCom.Parameters.Clear()
            MCom.Parameters.AddWithValue("@DriverID", DriverID)
            MCom.Parameters.AddWithValue("@DriverName", StrSQLRemoveChar(DriverName))
            MCom.Parameters.AddWithValue("@DriverPhone", DriverPhone)
            MCom.Parameters.AddWithValue("@VehicleNo", VehicleNo)
            MCom.Parameters.AddWithValue("@TrackingUrl", TrackingUrl)
            MCom.Parameters.AddWithValue("@ThirdParty", ThirdParty)
            MCom.Parameters.AddWithValue("@TrackNum", TrackNum)

            If MCom.ExecuteNonQuery = False Then
                Result(1) = "Gagal Update Live Tracking"
                GoTo Skip
            End If

            Result(0) = "0"
            Result(1) = ResponseOK
            Result(2) = ""

Skip:

            CreateResponseLog(wsAppName, wsAppVersion, Method, "SYS", Result, , TrackNum)

        Catch ex As Exception

            Dim e As New ClsError
            e.ErrorLog(wsAppName, wsAppVersion, Method, "", ex, Query)

        End Try

        Return Result

    End Function

    <WebMethod()>
    Public Function SetOtherExpeditionRTS(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        'untuk insert OtherExpeditionExpense, pada awb ekspedisi otomatis yang berstatus RTS

        Dim Method As String = "SetOtherExpeditionRTS"

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection

        Dim Query As String = ""

        Dim LogKeyword As String = Format(Date.Now, "yyMMddHHmmss")

        Dim User As String = Param(0).ToString
        Try
            Dim Password As String = Param(1).ToString
            Dim TrackNum As String = Param(2).ToString.Trim.ToUpper

            Dim oExpedition As String = ""
            Try
                oExpedition = Param(3).ToString.Trim.ToUpper
            Catch ex As Exception
                oExpedition = ""
            End Try
            Dim oAWB As String = ""
            Try
                oAWB = Param(4).ToString.Trim.ToUpper
            Catch ex As Exception
                oAWB = ""
            End Try
            Dim oWeight As String = ""
            Try
                oWeight = Param(5).ToString.Trim.ToUpper
            Catch ex As Exception
                oWeight = ""
            End Try
            Dim oAWBDate As String = ""
            Try
                oAWBDate = Param(6).ToString.Trim.ToUpper
            Catch ex As Exception
                oAWBDate = ""
            End Try

            Dim oAWBDCost As Double = -1
            Try
                If Not Param(7) Is Nothing Then
                    If IsNumeric(Param(7)) Then
                        oAWBDCost = CDbl(Param(7))
                        If oAWBDCost = 0 Then
                            oAWBDCost = -1
                        End If
                    End If
                End If
            Catch ex As Exception
                oAWBDCost = -1
            End Try

            Dim UserId As String = ""
            Try
                UserId = Param(8).ToString.Trim.ToUpper
            Catch ex As Exception
                UserId = ""
            End Try
            If UserId = "" Then
                UserId = User & "_RTS"
            End If


            Dim MTrn As MySqlTransaction

            MCon = MasterMCon.Clone

            Dim MCom As New MySqlCommand("", MCon)
            MCon.Open()

            LogKeyword = TrackNum & " " & oAWB

            CreateRequestLog(AppName, AppVersion, Method, User, "", Param, LogKeyword)

            MTrn = MCon.BeginTransaction()
            MCom.Transaction = MTrn

            Dim Process As Boolean = False

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            If UserOK(0) <> "0" Then
                Result(1) = UserOK(1)
                GoTo Skip
            End If

            If False Then
                If TrackNum = "" Then 'tidak perlu validasi awb ipp
                    Result(1) = "Input TrackNum"
                    GoTo Skip
                End If
            End If

            If oExpedition = "" Then
                Result(1) = "Input Ekspedisi"
                GoTo Skip
            End If

            If oAWB = "" Then
                Result(1) = "Input AWB 3PL"
                GoTo Skip
            End If

            If oAWBDCost < 0 Then
                Result(1) = "Input Biaya 3PL"
                GoTo Skip
            End If


            Dim MDa As MySqlDataAdapter = Nothing

            Dim dtRtsData As New DataTable

            MCom.Parameters.Clear()

            Query = "Select * From ExpAutoOrderTrx"
            Query &= " Where ThirdParty = @oExpedition And DeliveryId = @oAWB"
            If TrackNum <> "" Then
                Query &= " And TrackNum = @TrackNum"
                MCom.Parameters.AddWithValue("@TrackNum", TrackNum)
            End If
            MCom.CommandText = Query

            MCom.Parameters.AddWithValue("@oExpedition", oExpedition)
            MCom.Parameters.AddWithValue("@oAWB", oAWB)

            MDa = New MySqlDataAdapter(MCom)
            MDa.Fill(dtRtsData)

            If dtRtsData Is Nothing Then
                Result(1) = "Gagal query dtRtsData " & oAWB
                GoTo Skip
            End If
            If dtRtsData.Rows.Count < 1 Then
                Result(1) = "dtRtsData " & oAWB & " tidak ditemukan"
                GoTo Skip
            End If

            If TrackNum = "" Then
                TrackNum = dtRtsData.Rows(0).Item("TrackNum").ToString.Trim.ToUpper
            End If

            Query = "Select cast(concat(d.SjNum,'|',d.DriverId,'|',h.Code,'|',ifnull(b.Alias,h.Code)) as char) From SuratJalanD d"
            Query &= " Inner Join SuratJalanH h on (h.SjNum = d.SjNum and h.DriverId = d.DriverId)"
            Query &= " Left Join MstHub b on (b.Hub = h.Code)"
            Query &= " Where d.TrackNum = @TrackNum And d.`Status` = '0'"
            Query &= " Order By d.UpdTime Desc Limit 1"
            MCom.CommandText = Query

            MCom.Parameters.Clear()
            MCom.Parameters.AddWithValue("@TrackNum", TrackNum)

            Dim SjIppInfo As String = ""
            Try
                SjIppInfo = ("" & MCom.ExecuteScalar).ToString
            Catch ex As Exception
                SjIppInfo = ""
            End Try

            Dim SjIPP As String = ""
            Try
                SjIPP = SjIppInfo.Split("|")(0)
            Catch ex As Exception
                SjIPP = ""
            End Try

            Dim DriverIPP As String = ""
            Try
                DriverIPP = SjIppInfo.Split("|")(1)
            Catch ex As Exception
                DriverIPP = ""
            End Try

            Dim HubIPP As String = ""
            Try
                HubIPP = SjIppInfo.Split("|")(2)
            Catch ex As Exception
                HubIPP = ""
            End Try

            Dim HubIPPName As String = ""
            Try
                HubIPPName = SjIppInfo.Split("|")(3)
            Catch ex As Exception
                HubIPPName = ""
            End Try

            If SjIPP = "" Then
                Result(1) = "Data SJ " & TrackNum & " tidak ditemukan"
                GoTo Skip
            End If


            Dim eDstCity As Integer = 0

            Query = "Select DstCity From SuratJalanD Where SjNum = @SjIPP And TrackNum = TrackNum"
            MCom.CommandText = Query

            MCom.Parameters.Clear()
            MCom.Parameters.AddWithValue("@TrackNum", TrackNum)
            MCom.Parameters.AddWithValue("@SjIPP", SjIPP)

            Try
                eDstCity = MCom.ExecuteScalar
                If IsNumeric(eDstCity) = False Then
                    eDstCity = -1
                End If
            Catch ex As Exception
                eDstCity = -1
            End Try
            If eDstCity < 0 Then
                eDstCity = 0
            End If

            If eDstCity = 0 Then
                Query = "Select CoCityCode From Transaction Where TrackNum = @TrackNum"
                MCom.CommandText = Query

                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@TrackNum", TrackNum)

                Try
                    eDstCity = MCom.ExecuteScalar
                    If IsNumeric(eDstCity) = False Then
                        eDstCity = -1
                    End If
                Catch ex As Exception
                    eDstCity = -1
                End Try
                If eDstCity < 0 Then
                    eDstCity = 0
                End If
            End If

            If eDstCity = 0 Then
                Result(1) = "Gagal menentukan Kode Kota Tujuan " & TrackNum
                GoTo Skip
            End If

            Query = "Select Name From MstCity Where City = @eDstCity"
            MCom.CommandText = Query
            MCom.Parameters.Clear()
            MCom.Parameters.AddWithValue("@eDstCity", eDstCity)
            Dim eDstCityName As String = MCom.ExecuteScalar.ToString.Trim.ToUpper


            If oAWBDate = "" Or (IsDate(oAWBDate) = False) Then
                Result(1) = "Input Tanggal AWB dengan benar"
                GoTo Skip
            End If


            Dim AWB3PLBCKDTE As Integer = -1
            Try
                Query = "Select `Value` From Const Where Key1 = 'AWB3PLBCKDTE'"
                MCom.CommandText = Query
                MCom.Parameters.Clear()
                Dim TempAWB3PLBCKDTE As String = MCom.ExecuteScalar.ToString.Trim
                If TempAWB3PLBCKDTE = "" Or (IsNumeric(TempAWB3PLBCKDTE) = False) Then
                    AWB3PLBCKDTE = -1
                Else
                    AWB3PLBCKDTE = CInt(TempAWB3PLBCKDTE)
                End If
            Catch ex As Exception
                AWB3PLBCKDTE = -1
            End Try
            If AWB3PLBCKDTE = -1 Then
                AWB3PLBCKDTE = 3
            End If

            Query = "select cast(curdate() as char)"
            MCom.CommandText = Query
            MCom.Parameters.Clear()
            Dim SqlCurDate As String = MCom.ExecuteScalar

            'jagaan tanggal awb 3pl
            Query = "Select case when date_add('" & Format(CDate(SqlCurDate), "yyyy-MM-dd") & "', interval -" & AWB3PLBCKDTE & " day) <= '" & Format(CDate(oAWBDate), "yyyy-MM-dd") & "' then '1' else '0' end"
            MCom.CommandText = Query
            MCom.Parameters.Clear()
            If MCom.ExecuteScalar < 1 Then
                'If AutoAdjustExpAwbDate Then
                If True Then
                    oAWBDate = Format(DateAdd(DateInterval.Day, -1 * AWB3PLBCKDTE, CDate(SqlCurDate)), "yyyy-MM-dd")
                Else
                    Result(1) = "Tanggal AWB 3PL max. " & AWB3PLBCKDTE & " hari yang lalu (" & Format(DateAdd(DateInterval.Day, -1 * AWB3PLBCKDTE, CDate(SqlCurDate)), "yyyy-MM-dd") & ")"
                    GoTo Skip
                End If
            End If


            'validasi berat paket 3pl
            If oWeight <> "" And oExpedition <> "" Then

                Query = "Select ifnull(`Value`,0) From Const"
                Query &= " Where Key1 = 'MaxWeight3PL' And Key2 in (@oExpedition, '', '*')" 'dalam Gram
                Query &= " and curdate() between ActiveDate And ifnull(InActiveDate,curdate())"
                Query &= " order by (case when Key2 in ('', '*') then 9 else 1 end)"
                Query &= " limit 1"
                MCom.CommandText = Query

                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@oExpedition", oExpedition)

                Dim MaxInputWeight As Double = 0
                Try
                    MaxInputWeight = Math.Round(CDbl(MCom.ExecuteScalar) / 1000, 0) 'dalam KG
                Catch ex As Exception
                    MaxInputWeight = 0
                End Try

                Dim oExpAddInfo As String = ""
                Query = "Select AddInfo From Account Where Account = @oExpedition"
                MCom.CommandText = Query

                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@oExpedition", oExpedition)

                Try
                    oExpAddInfo = ("" & MCom.ExecuteScalar).ToString.Trim.ToUpper
                Catch ex As Exception
                    oExpAddInfo = ""
                End Try

                If oExpAddInfo.Contains("RNDWGTDCM3=Y") Then
                    oWeight = (Math.Round(CDbl(oWeight), 3, MidpointRounding.AwayFromZero)).ToString
                ElseIf oExpAddInfo.Contains("RNDWGTDCM1=Y") Then
                    oWeight = (Math.Round(CDbl(oWeight), 1, MidpointRounding.AwayFromZero)).ToString
                Else

                End If

                'If oWeight > MaxInputWeight Then
                '    Result(1) = "Berat max. " & MaxInputWeight & " KG"
                '    GoTo Skip
                'End If

            End If 'dari If oWeight <> "" And oExpedition <> ""

            If oAWB <> "" And oExpedition <> "" Then

                'validasi nomor resi berulang
                Query = "Select count(*) From OtherExpeditionExpense"
                Query &= " Where Expedition = '" & oExpedition & "'"
                Query &= " And ExpenseType = 'AWB'" 'ExpenseType disamakan dengan Hub Antar Alamat
                If False Then
                    Query &= " And replace( replace( replace(      AWB     ,'-','') ,'.',''), ' ','')"
                    Query &= "   = replace( replace( replace('" & oAWB & "','-','') ,'.',''), ' ','')"
                Else
                    Query &= " And AWB = '" & oAWB & "'"
                End If
                '2022-06-14 10:11, By Cucun, Jika Informasi nya sama persis boleh lewat
                '2022-06-16 09:39, By Cucun, Tidak disetujui cara seperti dibawah, langsung hubungi pa alex
                'Query &= " And ("
                'Query &= "   Ori <> '" & oOri & "'"
                'Query &= "   Or Dst <> '" & oDst & "'"
                'Query &= "   Or Weight <> '" & oWeight & "'"
                'Query &= " )"
                MCom.CommandText = Query
                MCom.Parameters.Clear()
                If MCom.ExecuteScalar > 0 Then
                    'tidak perlu dianggap error untuk RSC
                    'Result(1) = "Sudah pernah ada AWB 3PL (" & oExpedition & "." & oAWB & ")"
                    'GoTo Skip
                End If

            End If 'dari If oAWB <> "" And oExpedition <> ""


            Dim nMaxLengthOtherAWB As Integer = 300

            If oAWB <> "" And oAWB.Length > nMaxLengthOtherAWB Then
                Result(1) = "Isi nomor AWB 3PL dengan benar (max. " & nMaxLengthOtherAWB & " karakter)"
                GoTo Skip
            End If


            Query = "Update SuratJalanD"
            Query &= " Set OtherExpedition = @oExpedition, OtherExpeditionAWB = @oAWB, OtherExpeditionWeight = @oWeight"
            Query &= " , OtherExpeditionOri = @HubIPP, OtherExpeditionDst = @eDstCity"
            Query &= " , OtherExpeditionCost = @oAWBDCost"
            Query &= " , OtherExpeditionUpdTime = now(), OtherExpeditionUpdUser = @UserId"
            Query &= " , `Status` = '9'"
            Query &= " Where TrackNum = @TrackNum And SjNum = @SjIPP"
            Query &= " And `Status` = '0'"
            MCom.CommandText = Query

            MCom.Parameters.Clear()
            MCom.Parameters.AddWithValue("@TrackNum", TrackNum)
            MCom.Parameters.AddWithValue("@SjIPP", SjIPP)
            MCom.Parameters.AddWithValue("oExpedition", oExpedition)
            MCom.Parameters.AddWithValue("@oAWB", oAWB)
            MCom.Parameters.AddWithValue("@oWeight", oWeight)
            MCom.Parameters.AddWithValue("@HubIPP", HubIPP)
            MCom.Parameters.AddWithValue("@eDstCity", eDstCity)
            MCom.Parameters.AddWithValue("@oAWBDCost", oAWBDCost)
            MCom.Parameters.AddWithValue("@UserId", UserId)

            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(MCon, AppName, AppVersion, Method, "", ex, Query, LogKeyword)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try


            'catat biaya ekspedisi
            Query = "Insert ignore Into OtherExpeditionExpense ("
            Query &= " Expedition, AWB, ExpenseType"
            Query &= " , OriType, Ori, OriName, DstType, Dst, DstName"
            Query &= " , Weight, AWBDate, AddInfo"
            Query &= " , Cost, CostDraft, Rate, LeadTime, MinWeight, MinRate"
            Query &= " , AddTime, AddUser"
            Query &= " ) values ("
            Query &= " @oExpedition, @oAWB, 'AWB'"
            Query &= " , '1', @OriCode, @OriName, '2', @DstCode, @DstName"
            Query &= " , @oWeight, @oAWBDate, ''"
            Query &= " , @Cost, @CostDraft, @Rate, @LeadTime, 1, @Rate"
            Query &= " , now(), @UserId"
            Query &= " )"
            MCom.CommandText = Query

            MCom.Parameters.Clear()
            MCom.Parameters.AddWithValue("@oExpedition", oExpedition)
            MCom.Parameters.AddWithValue("@oAWB", oAWB)
            MCom.Parameters.AddWithValue("@OriCode", HubIPP)
            MCom.Parameters.AddWithValue("@OriName", HubIPPName)
            MCom.Parameters.AddWithValue("@DstCode", eDstCity)
            MCom.Parameters.AddWithValue("@DstName", eDstCityName)
            MCom.Parameters.AddWithValue("@oWeight", oWeight)
            MCom.Parameters.AddWithValue("@oAWBDate", oAWBDate)
            MCom.Parameters.AddWithValue("@LeadTime", "1")
            MCom.Parameters.AddWithValue("@Cost", oAWBDCost)
            MCom.Parameters.AddWithValue("@CostDraft", oAWBDCost)
            MCom.Parameters.AddWithValue("@Rate", oAWBDCost)
            MCom.Parameters.AddWithValue("@UserId", UserId)

            Try
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(MCon, AppName, AppVersion, Method, "", ex, Query, LogKeyword)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try


            Process = True

Skip:

            If Process Then
                MTrn.Commit()

                Result(0) = "0"
                Result(1) = ResponseOK
            Else
                MTrn.Rollback()
            End If

            CreateResponseLog(AppName, AppVersion, Method, User, Result, , LogKeyword)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Public Function GenerateDeliveryInfoPIN() As String()

        Dim Result(4) As String

        Dim ObjSQL As New ClsSQL

        Dim Template As String = "" 'pola / format PIN untuk proses dari Deliman kepada Toko Indomaret

        Dim nTry As Integer = 0
        Dim MaxTry As Integer = 5

        Dim MCon As MySqlConnection

        Try
            MCon = MasterMCon.Clone
            MCon.Open()

UlangTemplate:

            Template = ""

            Dim QueryPIN As String = ""

            QueryPIN &= " Select cast(concat("
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1') "
            QueryPIN &= "  ) as char) as PIN"

            QueryPIN &= "  , cast(concat("
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1') "
            QueryPIN &= "  ) as char) as PIN2"

            QueryPIN &= "  , cast(concat("
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1') "
            QueryPIN &= "  ) as char) as PIN3"

            QueryPIN &= "  , cast(concat("
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1') "
            QueryPIN &= " ) as char) as PIN4"

            QueryPIN &= "  , cast(concat("
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1') "
            QueryPIN &= "  ) as char) as PIN5"

            QueryPIN &= "  , cast(concat("
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1') "
            QueryPIN &= "  ) as char) as PIN6"

            QueryPIN &= "  , cast(concat("
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1'),"
            QueryPIN &= "  substring('12334567890',floor(rand()*10)+1,'1') "
            QueryPIN &= "  ) as char) as PIN7"

            Dim dtPINTemplate As DataTable = Nothing
            Try
                dtPINTemplate = ObjSQL.ExecDatatable(MCon, QueryPIN)
            Catch ex As Exception
                dtPINTemplate = Nothing
            End Try

            If dtPINTemplate Is Nothing Then
                dtPINTemplate = New DataTable
            End If

            'validasi PIN yang unik
            'If ObjSQL.ExecScalar(MCon, "Select count(BasePIN) From TransactionDeliveryInfo Where ifnull(BasePIN,'') = '" & Template & "'") > 0 Then
            'If ObjSQL.ExecScalar(MCon, "Select count(BasePIN) From TransactionDeliveryInfo Where BasePIN = '" & Template & "'") > 0 Then 'pengunaan IfNull menghilangkan fungsi index dan menyebabkan lemot
            '    If nTry < MaxTry Then
            '        nTry += 1
            '        GoTo UlangTemplate
            '    Else
            '        Template = "XXXXX"
            '    End If
            'End If

            If dtPINTemplate.Rows.Count > 0 Then

                Template = dtPINTemplate.Rows(0).Item("PIN").ToString

                If ObjSQL.ExecScalar(MCon, "Select count(BasePIN) From TransactionDeliveryInfo Where BasePIN = '" & Template & "'") > 0 Then

                    Template = dtPINTemplate.Rows(0).Item("PIN2").ToString

                    If ObjSQL.ExecScalar(MCon, "Select count(BasePIN) From TransactionDeliveryInfo Where BasePIN = '" & Template & "'") > 0 Then

                        Template = dtPINTemplate.Rows(0).Item("PIN3").ToString

                        If ObjSQL.ExecScalar(MCon, "Select count(BasePIN) From TransactionDeliveryInfo Where BasePIN = '" & Template & "'") > 0 Then

                            Template = dtPINTemplate.Rows(0).Item("PIN4").ToString

                            If ObjSQL.ExecScalar(MCon, "Select count(BasePIN) From TransactionDeliveryInfo Where BasePIN = '" & Template & "'") > 0 Then

                                Template = dtPINTemplate.Rows(0).Item("PIN5").ToString

                                If ObjSQL.ExecScalar(MCon, "Select count(BasePIN) From TransactionDeliveryInfo Where BasePIN = '" & Template & "'") > 0 Then

                                    Template = dtPINTemplate.Rows(0).Item("PIN6").ToString

                                    If ObjSQL.ExecScalar(MCon, "Select count(BasePIN) From TransactionDeliveryInfo Where BasePIN = '" & Template & "'") > 0 Then

                                        Template = dtPINTemplate.Rows(0).Item("PIN7").ToString

                                        If ObjSQL.ExecScalar(MCon, "Select count(BasePIN) From TransactionDeliveryInfo Where BasePIN = '" & Template & "'") > 0 Then

                                            Dim FinalTry As Boolean = False

                                            Dim Query As String = ""
                                            Query &= " Select cast(lpad(BasePin + 1,5,'0') as char) as PIN From TransactionDeliveryInfo"
                                            Query &= " Inner Join (select cast(substring('12334567890',floor(rand()*10)+1,'1') as char) as Prefix) x"
                                            Query &= " Where BasePin is not null And BasePin like concat(x.Prefix,'%')"
                                            Query &= " Order By BasePin Limit 5000"
                                            Dim dtFinalTry As DataTable = ObjSQL.ExecDatatable(MCon, Query)
                                            If Not dtFinalTry Is Nothing Then
                                                If dtFinalTry.Rows.Count > 0 Then

                                                    For f As Integer = 0 To dtFinalTry.Rows.Count - 1
                                                        Template = dtFinalTry.Rows(f).Item("PIN").ToString
                                                        If ObjSQL.ExecScalar(MCon, "Select count(BasePIN) From TransactionDeliveryInfo Where BasePIN = '" & Template & "'") > 0 Then
                                                            'continue loop
                                                        Else
                                                            FinalTry = True
                                                            Exit For
                                                        End If
                                                    Next

                                                End If
                                            End If

                                            If FinalTry = False Then
                                                If nTry < MaxTry Then
                                                    nTry += 1
                                                    GoTo UlangTemplate
                                                Else
                                                    Template = ""
                                                End If
                                            End If ' dari If FinalTry = False

                                        End If 'dari If CekPIN7

                                    End If 'dari If CekPIN6

                                End If 'dari If CekPIN5

                            End If 'dari If CekPIN4

                        End If 'dari If CekPIN3

                    End If 'dari If CekPIN2

                End If 'dari If CekPIN

            End If 'dari If dtPINTemplate.Rows.Count > 0

        Catch ex As Exception
            Template = ""
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


        If Template = "" Then
            Template = "XXXXX"
        End If

        Dim Alphabet(4) As String
        Alphabet(0) = "A" 'Pickup
        Alphabet(1) = "B" 'Cancel
        Alphabet(2) = "T" 'Keep
        Alphabet(3) = "K" 'Return
        Alphabet(4) = "" 'Base

        'menyisipkan karakter penanda pada PIN, sesuai tipe proses yang akan dijalankan
        For i As Integer = 0 To Result.Length - 1
            Result(i) = ""

            'Dim Position As Integer = GetRandom(0, 6)
            Dim Position As Integer = 0 'selalu prefix di awal

            If Position = 0 Then
                Result(i) = Alphabet(i) & Template
            ElseIf Position = 6 Then
                Result(i) = Template & Alphabet(i)
            Else
                Result(i) = Strings.Left(Template, Position) & Alphabet(i) & Strings.Right(Template, Template.Length - Position)
            End If

        Next


        Return Result

    End Function

    Public Function SetDiffTimeZoneByPostalCode(ByVal PostalCode As String, ByRef LocalTimeZone As Integer) As Integer

        'asumsinya adalah jam yang dikirim partner adalah selalu jam lokal
        'maka data yang disimpan akan dikonversi ke TimeZone Server HO

        Dim Result As Integer = 0

        Try
            Dim MyTimeZone As Integer = 7 'TimeZone Server HO
            LocalTimeZone = LocalTimeZoneByPostalCode(PostalCode)

            Result = LocalTimeZone - MyTimeZone
            If Result < 0 Then
                Result = 0
            End If
        Catch ex As Exception
            Result = 0
        End Try

        Return Result

    End Function

    Public Function LocalTimeZoneByPostalCode(ByVal PostalCode As String) As Integer

        Dim Result As String = ""

        If PostalCode <> "" Then
            Dim MCon As New MySqlConnection
            Try
                MCon = MasterMCon.Clone
                MCon.Open()

                Dim Query As String = ""
                Query = "Select TimeZone From MstPostalCode"
                Query &= " Where Code = @PostalCode"
                Query &= " Order By TimeZone Desc"

                SqlParam = New Dictionary(Of String, String)
                SqlParam.Add("@PostalCode", PostalCode)

                Dim ObjSQL As New ClsSQL
                Result = ("" & ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam)).ToString.Trim

                If IsNumeric(Result) = False Then
                    Result = ""
                End If

            Catch ex As Exception
                Result = ""

            Finally
                If MCon.State <> ConnectionState.Closed Then
                    MCon.Close()
                End If
                Try
                    MCon.Dispose()
                Catch ex As Exception
                End Try
            End Try
        End If

        If Result = "" Then
            Result = "7"
        End If

        Return CInt(Result)

    End Function

    <WebMethod()>
    Public Sub ServiceDeliverEmail()
        Dim Mcon As New MySqlConnection
        Dim Query As String = ""

        Try
            Mcon = MasterMCon.Clone

            Dim objSQL As New ClsSQL

            Query = " Select l.User, l.Password"
            Query &= " , i.Code as KodeToko, i.Name as NamaToko, i.DcCode as KodeDc, i.DcName as NamaDc"
            Query &= " From IndomaretStore_DeliveryMan d"
            Query &= " Inner Join MstLogin l on (l.User = concat('MS',d.Code) and l.Type='DSH' and l.Code = '1000319')"
            Query &= " Inner Join IndomaretStore_For_Report i on (i.Code = d.Code)"
            Query &= " Where d.Status = '8'"

            Dim dtQuery As DataTable = objSQL.ExecDatatable(Mcon, Query)

            If Not dtQuery Is Nothing Then
                If dtQuery.Rows.Count > 0 Then

                    Dim MailTo As String = ConfigurationManager.AppSettings("EmailSpv")
                    Dim Subject As String = "Data Username dan Password [NO REPLY]"
                    Dim Body As String = ""

                    Body = "<br />"
                    Body &= "Yth. SPV DMS"
                    Body &= "<br />"
                    Body &= "Berikut merupakan Username dan Password untuk"
                    Body &= " Nama Toko,Kode Toko, Kode Dc dan Nama DC :"
                    Body &= "<Table border='1' style='border-collapse:collapse'>"
                    Body &= "<tr>"
                    Body &= "<td> Username </td>"
                    Body &= "<td> Password </td>"
                    Body &= "<td> Kode Toko </td>"
                    Body &= "<td> Nama Toko </td>"
                    Body &= "<td> Kode Dc </td>"
                    Body &= "<td> Nama Dc </td>"
                    Body &= "</tr>"

                    For i As Integer = 0 To dtQuery.Rows.Count - 1
                        Body &= "<tr>"
                        Body &= "<td>" & dtQuery.Rows(i).Item("User").ToString & "</td>"
                        Body &= "<td>" & dtQuery.Rows(i).Item("Password").ToString & "</td>"
                        Body &= "<td>" & dtQuery.Rows(i).Item("KodeToko").ToString & "</td>"
                        Body &= "<td>" & dtQuery.Rows(i).Item("NamaToko").ToString & "</td>"
                        Body &= "<td>" & dtQuery.Rows(i).Item("KodeDc").ToString & "</td>"
                        Body &= "<td>" & dtQuery.Rows(i).Item("NamaDC").ToString & "</td>"
                        Body &= "</tr>"
                    Next

                    Body &= "</table>"
                End If
            End If
            Try
                'EmailSend(MailTo, Subject, Body, "", _EmailType._DEFAULT)
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog("", "", "ServiceDeliverEmail", "", ex, Query)
        Finally
            If Mcon.State <> ConnectionState.Closed Then
                Mcon.Close()
            End If
            Try
                Mcon.Dispose()
            Catch ex As Exception
            End Try
        End Try
    End Sub

    Enum _EmailType
        _DEFAULT
        _PAKET
        _REPORT
        _COMPLAIN
        _HADIAHPOINKU
    End Enum
    'Private Function EmailSend(ByVal _MailTo As String, ByVal _Subject As String, ByVal _Body As String, ByVal _Filename As String _
    ', ByVal eType As _EmailType, Optional ByVal HandlingLabel As String = "", Optional ByVal SleepTime As Integer = 1500 _
    ', Optional ByVal BccIndopaket As Boolean = False, Optional ByVal _Filenames As String() = Nothing, Optional ByVal _MailReplyTo As String = "", Optional ByVal BccList As String = "") As Boolean 'kirim email

    '    Dim Method As String = "EmailSend"

    '    Dim MCon As New MySqlConnection

    '    Dim Query As String = ""

    '    Dim SqlParam As New Dictionary(Of String, String)

    '    '2020-09-24, ganti ke support@indopaket, kadang error sender unknown
    '    Dim _EmailFrom As String = "support@indopaket.co.id"
    '    Dim _EmailTo As String = ""
    '    Dim _EmailBCC As String = ""
    '    Dim _EmailDisplayName As String = "INDOPAKET"
    '    Dim _EmailReplyTo As String = ""
    '    Dim _AttachmentList As String = ""
    '    Dim _LogMaxColumnLength As Long = 767

    '    Dim _EmailUsername As String = "support@indopaket.co.id"
    '    Dim _EmailPassword As String = "support@ipp18"
    '    Dim _EmailServer As String = "172.20.20.85"
    '    Dim _EmailPort As Integer = 25

    '    Try
    '        MCon = MasterMCon.Clone
    '        MCon.Open()

    '        If _MailTo.Trim <> "" Then

    '            '=== Setup Email
    '            Dim ConfigEmailServer As String = ""
    '            Try
    '                ConfigEmailServer = "" & ConfigurationManager.AppSettings("EmailServerIndopaket")
    '            Catch ex As Exception
    '                ConfigEmailServer = ""
    '            End Try

    '            If ConfigEmailServer = "" Then
    '                ConfigEmailServer = "172.20.20.85"
    '            End If

    '            Dim TracelogEmail As String = "0"
    '            Try
    '                TracelogEmail = "" & ConfigurationManager.AppSettings("TracelogEmail")
    '                If TracelogEmail = "" Then
    '                    TracelogEmail = "0"
    '                End If
    '            Catch ex As Exception
    '                TracelogEmail = "0"
    '            End Try

    '            Select Case eType
    '                Case _EmailType._PAKET
    '                    _EmailFrom = "info@indopaket.co.id"
    '                    _EmailUsername = "info@indopaket.co.id"
    '                    _EmailPassword = "info@ipp18"
    '                    _EmailServer = ConfigEmailServer
    '                    _EmailPort = 25
    '                    _EmailDisplayName = "INDOPAKET"

    '                    _EmailReplyTo = "support@indopaket.co.id"

    '                Case _EmailType._REPORT
    '                    _EmailFrom = "report@indopaket.co.id"
    '                    _EmailUsername = "report@indopaket.co.id"
    '                    _EmailPassword = "report@ipp18"
    '                    _EmailServer = ConfigEmailServer
    '                    _EmailPort = 25
    '                    _EmailDisplayName = "INDOPAKETReport"

    '                Case _EmailType._COMPLAIN
    '                    _EmailFrom = "info@indopaket.co.id"
    '                    _EmailUsername = "info@indopaket.co.id"
    '                    _EmailPassword = "info@ipp18"
    '                    _EmailServer = ConfigEmailServer
    '                    _EmailPort = 25
    '                    _EmailDisplayName = _MailReplyTo

    '                Case _EmailType._HADIAHPOINKU
    '                    _EmailFrom = "info@indopaket.co.id"
    '                    _EmailUsername = "info@indopaket.co.id"
    '                    _EmailPassword = "info@ipp18"
    '                    _EmailServer = ConfigEmailServer
    '                    _EmailPort = 25
    '                    _EmailDisplayName = "INDOPAKET Hadiah POINKU"

    '                    _EmailReplyTo = "hadiahpoinku@indopaket.co.id"
    '            End Select

    '            If _MailReplyTo <> "" Then
    '                _EmailReplyTo = _MailReplyTo
    '            End If

    '            If StringTest <> "" Then
    '                _EmailDisplayName &= "DEV"
    '            End If

    '            If TracelogEmail <> "0" Then
    '                Dim e As New ClsError
    '                e.DebugLog(MCon, "Email", "", "TracelogEmail", "", "Email Server " & _EmailServer)
    '            End If

    '            _EmailTo = _MailTo

    '            _EmailBCC = ""
    '            'If BccIndopaket Then
    '            '    _EmailBcc = "budil@indopaket.co.id;cucun@indopaket.co.id;alvinadi.widjaja@indopaket.co.id"
    '            'End If
    '            If BccList <> "" Then
    '                If _EmailBCC <> "" Then
    '                    _EmailBCC &= ";"
    '                End If
    '                _EmailBCC &= BccList
    '            End If

    '            Dim MyMailMessage As New MailMessage()

    '            MyMailMessage.From = New MailAddress(_EmailFrom, _EmailDisplayName)
    '            MyMailMessage.Sender = New MailAddress(_EmailFrom, _EmailDisplayName)

    '            Dim dtEmailList As New DataTable
    '            dtEmailList.Columns.Add("Email")

    '            If _EmailTo <> "" Then

    '                Dim ListTo As String() = _EmailTo.Split(";")
    '                For i As Integer = 0 To ListTo.Length - 1
    '                    If ListTo(i) <> "" Then
    '                        Dim ListTo2 As String() = ListTo(i).Split(",")
    '                        For j As Integer = 0 To ListTo2.Length - 1
    '                            If ListTo2(j) <> "" Then
    '                                dtEmailList.Rows.Add(New String() {ListTo2(j)})
    '                            End If
    '                        Next
    '                    End If
    '                Next

    '                If dtEmailList.Rows.Count > 0 Then
    '                    For i As Integer = 0 To dtEmailList.Rows.Count - 1
    '                        MyMailMessage.To.Add(dtEmailList.Rows(i).Item("Email").ToString)
    '                    Next
    '                End If

    '            End If


    '            If _EmailBCC <> "" Then

    '                dtEmailList.Rows.Clear()

    '                Dim ListBcc As String() = _EmailBCC.Split(";")
    '                For i As Integer = 0 To ListBcc.Length - 1
    '                    If ListBcc(i) <> "" Then
    '                        Dim ListBcc2 As String() = ListBcc(i).Split(",")
    '                        For j As Integer = 0 To ListBcc2.Length - 1
    '                            If ListBcc2(j) <> "" Then
    '                                dtEmailList.Rows.Add(New String() {ListBcc2(j)})
    '                            End If
    '                        Next
    '                    End If
    '                Next

    '                If dtEmailList.Rows.Count > 0 Then
    '                    For i As Integer = 0 To dtEmailList.Rows.Count - 1
    '                        MyMailMessage.Bcc.Add(dtEmailList.Rows(i).Item("Email").ToString)
    '                    Next
    '                End If

    '            End If

    '            If _EmailReplyTo <> "" Then
    '                Dim ReplyTo As MailAddress = New MailAddress(_EmailReplyTo, _EmailDisplayName)
    '                MyMailMessage.ReplyTo = ReplyTo
    '            End If

    '            MyMailMessage.Subject = _Subject
    '            If StringTest <> "" Then
    '                MyMailMessage.Subject &= " (HANYA UNTUK KEPERLUAN DEVELOPMENT / TESTING)"
    '            End If
    '            MyMailMessage.Body = _Body
    '            MyMailMessage.IsBodyHtml = True

    '            Dim Attach As System.Net.Mail.Attachment = Nothing
    '            If _Filename.Trim <> "" Then
    '                If IO.File.Exists(_Filename) Then
    '                    Attach = New System.Net.Mail.Attachment(_Filename)
    '                    MyMailMessage.Attachments.Add(Attach)
    '                End If
    '            End If

    '            'If AttachHandlingLabel Then
    '            '    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\awayheat.jpg"))
    '            '    MyMailMessage.Attachments.Add(Attach)
    '            '    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\fragile.jpg"))
    '            '    MyMailMessage.Attachments.Add(Attach)
    '            '    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\heavy.jpg"))
    '            '    MyMailMessage.Attachments.Add(Attach)
    '            '    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\keepdry.jpg"))
    '            '    MyMailMessage.Attachments.Add(Attach)
    '            '    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\thiswayup.jpg"))
    '            '    MyMailMessage.Attachments.Add(Attach)
    '            'End If

    '            If HandlingLabel <> "" Then
    '                If HandlingLabel.Contains("HEAT") Then
    '                    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\awayheat.jpg"))
    '                    MyMailMessage.Attachments.Add(Attach)
    '                End If
    '                If HandlingLabel.Contains("FRAGILE") Then
    '                    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\fragile.jpg"))
    '                    MyMailMessage.Attachments.Add(Attach)
    '                End If
    '                If HandlingLabel.Contains("HEAVY") Then
    '                    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\heavy.jpg"))
    '                    MyMailMessage.Attachments.Add(Attach)
    '                End If
    '                If HandlingLabel.Contains("DRY") Then
    '                    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\keepdry.jpg"))
    '                    MyMailMessage.Attachments.Add(Attach)
    '                End If
    '                If HandlingLabel.Contains("UPRIGHT") Then
    '                    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\thiswayup.jpg"))
    '                    MyMailMessage.Attachments.Add(Attach)
    '                End If
    '            End If

    '            If _Filenames Is Nothing = False Then
    '                For i As Integer = 0 To _Filenames.Length - 1
    '                    If ("" & _Filenames(i)).Trim <> "" Then
    '                        If IO.File.Exists(_Filenames(i)) Then
    '                            Attach = New System.Net.Mail.Attachment(_Filenames(i))
    '                            MyMailMessage.Attachments.Add(Attach)

    '                            If _AttachmentList <> "" Then
    '                                _AttachmentList &= "|"
    '                            End If
    '                            _AttachmentList &= _Filenames(i)
    '                        End If
    '                    End If
    '                Next
    '            End If

    '            ''=== Setup POP3
    '            'Dim email As Pop3Client
    '            'Try
    '            '    email = New Pop3Client(_EmailUsername, _EmailPassword, _EmailServer)
    '            '    email.OpenInbox()
    '            'Catch ex As Exception
    '            'End Try


    '            '=== Setup SMTP

    '            Dim SMTPServer As New SmtpClient(_EmailServer)

    '            SMTPServer.Port = _EmailPort

    '            SMTPServer.Credentials = New System.Net.NetworkCredential(_EmailUsername, _EmailPassword)

    '            SMTPServer.Send(MyMailMessage)

    '            System.Threading.Thread.Sleep(SleepTime)

    '            SMTPServer = Nothing

    '            MyMailMessage = Nothing

    '            If Not Attach Is Nothing Then
    '                Try
    '                    Attach.Dispose()
    '                Catch ex As Exception
    '                End Try
    '            End If

    '            'catat log

    '            Dim ObjSQL As New ClsSQL

    '            Query = " Insert Into emaillog ("
    '            Query &= " `subject`, `body`, `attachment`, `emailfrom`, `emailto`, `emailbcc`, `emailreplyto`, `status`, `description`, `updtime`, `upduser`"
    '            Query &= " ) values ("
    '            Query &= " @_Subject, @_Body, @_Attachment, @_EmailFrom, @_EmailTo, @_EmailBcc, @_EmailReplyTo, @Status , @Description , now()    , @UpdUser"
    '            Query &= " )"

    '            SqlParam = New Dictionary(Of String, String)
    '            SqlParam.Add("@_Subject", Left(_Subject, _LogMaxColumnLength))
    '            SqlParam.Add("@_Body", _Body)
    '            SqlParam.Add("@_Attachment", _AttachmentList)
    '            SqlParam.Add("@_EmailFrom", Left(_EmailFrom, _LogMaxColumnLength))
    '            SqlParam.Add("@_EmailTo", Left(_EmailTo, _LogMaxColumnLength))
    '            SqlParam.Add("@_EmailBcc", Left(_EmailBCC, _LogMaxColumnLength))
    '            SqlParam.Add("@_EmailReplyTo", _EmailReplyTo)
    '            SqlParam.Add("@Status", "1")
    '            SqlParam.Add("@Description", "success")
    '            SqlParam.Add("@UpdUser", "SYS")

    '            ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

    '            'Try
    '            '    email.CloseConnection()
    '            'Catch ex As Exception
    '            'End Try

    '        Else

    '            Dim e As New ClsError
    '            e.DebugLog(MCon, "Email", "", Method, "", "Skip Email " & _Subject)

    '        End If

    '        Return True

    '    Catch ex As Exception

    '        Dim ObjSQL As New ClsSQL

    '        Query = " Insert Into emaillog ("
    '        Query &= " `subject`, `body`, `attachment`, `emailfrom`, `emailto`, `emailbcc`, `emailreplyto`, `status`, `description`, `updtime`, `upduser`"
    '        Query &= " ) values ("
    '        Query &= " @_Subject, @_Body, @_Attachment, @_EmailFrom, @_EmailTo, @_EmailBcc, @_EmailReplyTo, @Status , @Description , now()    , @UpdUser"
    '        Query &= " )"

    '        SqlParam = New Dictionary(Of String, String)
    '        SqlParam.Add("@_Subject", Left(_Subject, _LogMaxColumnLength))
    '        SqlParam.Add("@_Body", _Body)
    '        SqlParam.Add("@_Attachment", _AttachmentList)
    '        SqlParam.Add("@_EmailFrom", Left(_EmailFrom, _LogMaxColumnLength))
    '        SqlParam.Add("@_EmailTo", Left(_EmailTo, _LogMaxColumnLength))
    '        SqlParam.Add("@_EmailBcc", Left(_EmailBCC, _LogMaxColumnLength))
    '        SqlParam.Add("@_EmailReplyTo", _EmailReplyTo)
    '        SqlParam.Add("@Status", "9")
    '        SqlParam.Add("@Description", ex.Message)
    '        SqlParam.Add("@UpdUser", "SYS")

    '        ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

    '        Dim e As New ClsError
    '        e.ErrorLog("Email", "", Method, "", ex, "MailFrom : " & _EmailFrom & vbCrLf & "MailTo : " & _MailTo & vbCrLf & "Subject : " & _Subject)

    '        Return False

    '    Finally

    '        If MCon.State <> ConnectionState.Closed Then
    '            MCon.Close()
    '        End If
    '        Try
    '            MCon.Dispose()
    '        Catch ex As Exception
    '        End Try

    '    End Try

    'End Function

    Private Function NotificationWording(ByVal MCon As MySqlConnection, ByVal Type As String, ByVal Status As String, ByVal Account As String) As String

        Dim Result As String = ""

        Dim Query As String = ""
        Try
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            Query = "Select Wording From ("

            Query &= " Select '1' as OrderNo, cast(Wording as char) as Wording"
            Query &= " From NotificationWording"
            Query &= " Where `Type` = '" & Type & "' And `Status` = '" & Status & "'"
            Query &= " And Account = '" & Account & "'"

            Query &= " union"

            Query &= " Select '9' as OrderNo, cast(Wording as char) as Wording"
            Query &= " From NotificationWording"
            Query &= " Where `Type` = '" & Type & "' And `Status` = '" & Status & "'"
            Query &= " And Account = '" & "*" & "'" '-- default

            Query &= " )x"
            Query &= " Order By OrderNo"
            Query &= " Limit 1"

            Dim ObjSQL As New ClsSQL
            Result = ("" & ObjSQL.ExecScalar(MCon, Query)).ToString

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog("", "", "NotificationWording", "", ex, Query)

            Result = ""
        End Try

        Return Result

    End Function

    Private Function EmailFooter(ByVal MCon As MySqlConnection, Optional ByVal IncludeThanks As Boolean = True, Optional ByVal SuggestReply As Boolean = False) As String

        Dim Result As String = ""

        Dim Query As String = ""

        Query = "Select ServiceName, Website, CallCenter"
        Query &= " From `Company`"

        Dim ObjSQL As New ClsSQL
        Dim dtCompany As DataTable = ObjSQL.ExecDatatable(MCon, Query)

        Result = "<br />"

        If IncludeThanks Then
            Result &= "<br />"
            Result &= "Terima kasih telah menggunakan layanan " & dtCompany.Rows(0).Item("ServiceName").ToString.ToUpper
        End If

        If dtCompany.Rows(0).Item("Website").ToString <> "" Then
            Dim Website As String = dtCompany.Rows(0).Item("Website").ToString
            Website = Replace(Replace(Website, "http://", ""), "https://", "")

            Dim HttpWebsite As String = "http://" & Website

            Result &= "<br />"
            Result &= "W: <a href=""" & HttpWebsite & """ target=""_blank"">" & Website & "</a>"
        End If
        If dtCompany.Rows(0).Item("CallCenter").ToString <> "" Then
            Result &= "<br />"
            Result &= "C: " & dtCompany.Rows(0).Item("CallCenter").ToString
        End If

        If SuggestReply = False Then
            Result &= "<br />"
            Result &= "<br />"
            Result &= "<i>" & "Email ini dibuat secara otomatis, tidak perlu ditanggapi" & "</i>"
        End If

        Return Result

    End Function

    <WebMethod()>
    Public Function GetDCList(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            Dim DCCode As String = ""
            Try
                DCCode = Param(2).ToString
            Catch ex As Exception
                DCCode = ""
            End Try

            MCon = MasterMCon.Clone

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'Else

            Dim ObjSQL As New ClsSQL

            Query = "Select Code as Value"
            Query &= " , Cast(Concat(replace(Name,',',''), ' (', Code, ')') As Char(50)) As Display"
            Query &= " From indomaretdc Where 1=1"
            Query &= " And curdate() between ActiveDate and IfNull(InactiveDate,curdate())"
            If DCCode.Trim <> "" Then
                Query &= " And Code = '" & DCCode & "'"
            End If
            Query &= " Order By Name"

            Dim dtQuery As DataTable = ObjSQL.ExecDatatable(MCon, Query)

            If Not dtQuery Is Nothing Then

                Dim dtTemp As New DataTable

                dtTemp = dtQuery.Clone

                dtTemp.Rows.Add(New String() {"", ""})

                For Each drow As DataRow In dtQuery.Rows
                    dtTemp.ImportRow(drow)
                Next

                Result(2) = ConvertDatatableToString(dtTemp)
                Result(1) = ResponseOK
                Result(0) = "0"

            End If

            'End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetDCList", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Private Function GetStore(ByVal MCon As MySqlConnection, ByVal StoreCode As String _
    , Optional ByVal IsNew As Boolean = False, Optional ByVal ShAccount As String = "") As String

        Dim Store As String = ""

        Dim ObjSQL As New ClsSQL

        Try
            'Store = "" & ObjSQL.ExecScalar(MCon, "Select `Code` From IndomaretStore Where `Code` = '" & StoreCode & "'" & IIf(IsNew, " And FlgNewTrx = 'Y'", ""))

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@StoreCode", StoreCode)

            Store = "" & ObjSQL.ExecScalarWithParam(MCon, "Select `Code` From IndomaretStore Where `Code` = @StoreCode" & IIf(IsNew, " And FlgNewTrx = 'Y'", ""), SqlParam)

        Catch ex As Exception
            Store = ""
        End Try

        If Store <> "" Then
            If IsNew And ShAccount <> "" Then 'validasi toko sesuai coverage dc yang disepakati
                Try
                    'Store = "" & ObjSQL.ExecScalar(MCon, "Select Store as `Code` From MappingDCStore" & _
                    '" Where Store = '" & StoreCode & "' And DC in" & _
                    '" (Select DC From AccountDC Where `Account` = '" & ShAccount & "'" & _
                    '" And curdate() between ActiveDate and IfNull(InactiveDate,curdate()))")

                    SqlParam = New Dictionary(Of String, String)
                    SqlParam.Add("@StoreCode", StoreCode)
                    SqlParam.Add("@ShAccount", ShAccount)

                    Store = "" & ObjSQL.ExecScalarWithParam(MCon,
                    " Select Store as `Code` From MappingDCStore" &
                    " Where Store = @StoreCode And DC in" &
                    " (Select DC From AccountDC Where `Account` = @ShAccount" &
                    " And curdate() between ActiveDate and IfNull(InactiveDate,curdate()))", SqlParam)

                Catch ex As Exception
                    Store = ""
                End Try
            End If
        End If

        Return Store

    End Function

    Private Function GetStoreNotFound(ByVal MCon As MySqlConnection, ByVal StoreCode As String) As String

        'Dim Reason As String = "Select Reason From Indomaretstore_Deleted Where Code = '" & StoreCode & "'"
        Dim Reason As String = ""
        Try
            Dim ObjSQL As New ClsSQL
            'Reason = ("" & ObjSQL.ExecScalar(MCon, Reason)).ToString.Trim

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@StoreCode", StoreCode)

            Reason = ("" & ObjSQL.ExecScalarWithParam(MCon, "Select Reason From Indomaretstore_Deleted Where Code = @StoreCode", SqlParam)).ToString.Trim

        Catch ex As Exception
            Reason = ""
        End Try

        Return "Toko " & StoreCode & " tidak ditemukan / tidak aktif / tidak cover" & IIf(Reason <> "", " (" & Reason & ")", "")

    End Function

    Private ValidateStoreCheckTransactionDeliveryInfo As Boolean = False

    Public Function ConvertNumListToSQL(ByVal NumList As String) As String

        On Error Resume Next

        'convert ke versi SQL, supaya bisa pakai in (...)
        Dim NumListSQL As String = NumList.Trim

        NumListSQL = NumListSQL.Replace("'", "")
        While NumListSQL.Contains("| ")
            NumListSQL = NumListSQL.Replace("| ", "|")
        End While
        While NumListSQL.Contains(" |")
            NumListSQL = NumListSQL.Replace(" |", "|")
        End While
        If NumListSQL.StartsWith("|") Then
            NumListSQL = Strings.Right(NumListSQL, NumListSQL.Length - 1)
        End If
        If NumListSQL.EndsWith("|") Then
            NumListSQL = Strings.Left(NumListSQL, NumListSQL.Length - 1)
        End If
        NumListSQL = NumListSQL.Replace("|", "','")
        NumListSQL = "'" & NumListSQL & "'"

        Return NumListSQL

    End Function

    <WebMethod()>
    Public Function UpdateStatusDeliveryInfo() As Object()

        Dim AppName As String = "Web Report Dev"
        Dim AppVersion As String = "22.08.18.00"

        Dim Result As Object() = CreateResult()

        Dim Method As String = "UpdateStatusDeliveryInfo"

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim TrackStatus As String = ""

        Dim Param(15) As Object

        Param(0) = "2015548000"
        Param(1) = "2015548000"
        Param(2) = "2024-05-06 3:19:00"
        Param(3) = "PICKUP"
        Param(4) = "Indomaret"
        Param(5) = "STO"
        Param(6) = "T001"
        Param(7) = "2015548000"
        Param(8) = "Ryu"
        Param(9) = "IndoGrosir"
        Param(10) = "RTO"
        Param(11) = "T002"
        Param(12) = "2015548000"
        Param(13) = "Dudung"
        Param(14) = "BJX220512140329"
        Param(15) = "XX00"


        Dim User As String = Param(0).ToString 'user yang terdaftar sistem

        Try
            '=== request parameter
            Dim Password As String = Param(1).ToString

            Dim TrackDateTime As DateTime = CDate(Param(2)) 'waktu aktual tracking
            Dim TrackDate As String = Format(TrackDateTime, "yyyy-MM-dd")
            Dim TrackTime As String = Format(TrackDateTime, "HH:mm:ss")

            TrackStatus = Param(3).ToString.ToUpper

            'pihak 1 - pemberi
            Dim SenderCompany As String = Param(4).ToString.ToUpper.Trim 'kode perusahaan
            Dim SenderType As String = Param(5).ToString.ToUpper.Trim 'tipe operator, misal: 'STO' = Toko, dan lain2
            Dim SenderCode As String = Param(6).ToString.ToUpper.Trim 'kode operator, misal: 'T001', dan lain2
            Dim SenderID As String = Param(7).ToString.Trim.Replace("'", "").Replace(",", "").Replace("|", "") 'NIK operator
            Dim SenderName As String = Param(8).ToString.ToUpper.Trim.Replace("'", "").Replace(",", "").Replace("|", "") 'Nama operator
            If SenderName <> "" Then
                SenderName = Strings.Left(StrSQLRemoveChar(SenderName), 45)
            End If

            'pihak 2 - penerima
            Dim ReceiverCompany As String = Param(9).ToString.ToUpper.Trim
            Dim ReceiverType As String = Param(10).ToString.ToUpper.Trim
            Dim ReceiverCode As String = Param(11).ToString.ToUpper.Trim
            Dim ReceiverID As String = Param(12).ToString.Trim.Replace("'", "").Replace(",", "").Replace("|", "")
            Dim ReceiverName As String = Param(13).ToString.ToUpper.Trim.Replace("'", "").Replace(",", "").Replace("|", "")
            If ReceiverName <> "" Then
                ReceiverName = Strings.Left(StrSQLRemoveChar(ReceiverName), 45)
            End If

            Dim TrackNumList As String = Param(14).ToString.Trim.ToUpper 'nomor resi yang di-update status-nya
            Dim RefNumber As String = Param(15).ToString.Trim.ToUpper 'nomor referensi lain (bila ada), misal: nomor surat jalan toko, dan lain2

            '=== request parameter


            MCon = MasterMCon.Clone
            MCon.Open()

            Dim MCom As New MySqlCommand("", MCon)


            Dim LogKeyword As String = ""
            Try
                LogKeyword &= " " & SenderCode
            Catch ex As Exception
            End Try
            Try
                LogKeyword &= " " & ReceiverCode
            Catch ex As Exception
            End Try
            Try
                LogKeyword &= " " & TrackStatus
            Catch ex As Exception
            End Try
            Try
                LogKeyword &= " " & TrackNumList
            Catch ex As Exception
            End Try
            LogKeyword = Left(Trim(LogKeyword), 100)


            CreateRequestLog(AppName, AppVersion, Method & " " & TrackStatus, User, SenderCode & " " & ReceiverCode, Param, LogKeyword)

            Dim Process As Boolean = False


            '=== begin validasi

            'If TrackStatus.Trim = "" Then
            '    Result(1) = "Input Update Status"
            '    GoTo Skip
            'End If

            'If TrackNumList.Trim = "" Then
            '    Result(1) = "Input TrackNum"
            '    GoTo Skip
            'End If

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            '    Result(1) = UserOK(1)
            '    GoTo Skip
            'End If


            Select Case TrackStatus

                Case "PICKUP", "PICKUP_CANCEL" 'paket diambil di Toko oleh DeliveryMan

                    'Sender = Toko
                    'Receiver = DeliveryMan

                    If ("" & SenderCode) = "" Then
                        Result(1) = "Input Kode Toko"
                        GoTo Skip
                    End If

                    If ValidateStoreCheckTransactionDeliveryInfo Then
                        If GetStore(MCon, SenderCode) = "" Then
                            Result(1) = GetStoreNotFound(MCon, SenderCode)
                            GoTo Skip
                        End If
                    End If

                    If ReceiverName = "" Then
                        Result(1) = "Input Nama Penerima"
                        GoTo Skip
                    End If

                    Dim TrackNumAll As String() = TrackNumList.Split("|")

                    For i As Integer = 0 To TrackNumAll.Length - 1

                        'TrackNum,TrxPartner
                        Dim TrackNumSplit As String() = TrackNumAll(i).Split(",")

                        Dim TrackNum As String = ""
                        Try
                            TrackNum = TrackNumSplit(0).ToUpper
                        Catch ex As Exception
                            TrackNum = ""
                        End Try
                        Dim TrxPartner As String = ""
                        Try
                            TrxPartner = TrackNumSplit(1).ToUpper
                        Catch ex As Exception
                            TrxPartner = ""
                        End Try
                        If TrxPartner = "" Then
                            TrxPartner = TrackNum
                        End If

                        If TrackNum = "" Or TrxPartner = "" Then
                            Result(1) = "Input AWB dan TrxPartner"
                            GoTo Skip
                        End If

                        Query = "Select count(TrackNum)"
                        Query &= " From `transactiondeliveryinfo`"
                        Query &= " Where TrackNum = '" & TrackNum & "'"
                        MCom.CommandText = Query
                        If MCom.ExecuteScalar < 1 Then
                            Result(1) = TrackNum & " tidak ditemukan"
                            GoTo Skip
                        End If

                    Next


                    'dari Case "PICKUP"


                Case "RETURN" 'pengembalian bulky ke Toko oleh DeliveryMan

                    'Sender = DeliveryMan
                    'Receiver = Toko

                    If ("" & ReceiverCode) = "" Then
                        Result(1) = "Input Kode Toko"
                        GoTo Skip
                    End If

                    If ValidateStoreCheckTransactionDeliveryInfo Then
                        If GetStore(MCon, ReceiverCode) = "" Then
                            Result(1) = GetStoreNotFound(MCon, ReceiverCode)
                            GoTo Skip
                        End If
                    End If

                    If SenderName = "" Then
                        Result(1) = "Input Nama Pengirim"
                        GoTo Skip
                    End If

                    Dim TrackNumAll As String() = TrackNumList.Split("|")

                    For i As Integer = 0 To TrackNumAll.Length - 1

                        'TrackNum,TrxPartner
                        Dim TrackNumSplit As String() = TrackNumAll(i).Split(",")

                        Dim TrackNum As String = ""
                        Try
                            TrackNum = TrackNumSplit(0).ToUpper
                        Catch ex As Exception
                            TrackNum = ""
                        End Try
                        Dim TrxPartner As String = ""
                        Try
                            TrxPartner = TrackNumSplit(1).ToUpper
                        Catch ex As Exception
                            TrxPartner = ""
                        End Try

                        If TrackNum = "" Or TrxPartner = "" Then
                            Result(1) = "Input AWB dan TrxPartner"
                            GoTo Skip
                        End If

                        Query = "Select count(TrackNum)"
                        Query &= " From `TransactionDeliveryInfo`"
                        Query &= " Where TrackNum = '" & TrackNum & "'"
                        MCom.CommandText = Query
                        If MCom.ExecuteScalar < 1 Then
                            Result(1) = TrackNum & " tidak ditemukan"
                            GoTo Skip
                        End If

                    Next


                    'dari Case "RETURN"


                Case "CANCEL" 'pembatalan pengiriman oleh DeliveryMan

                    'Sender = DeliveryMan
                    'Receiver = Toko

                    If ("" & ReceiverCode) = "" Then
                        Result(1) = "Input Kode Toko"
                        GoTo Skip
                    End If

                    If ValidateStoreCheckTransactionDeliveryInfo Then
                        If GetStore(MCon, ReceiverCode) = "" Then
                            Result(1) = GetStoreNotFound(MCon, ReceiverCode)
                            GoTo Skip
                        End If
                    End If

                    If SenderName = "" Then
                        Result(1) = "Input Nama Pengirim"
                        GoTo Skip
                    End If

                    Dim TrackNumAll As String() = TrackNumList.Split("|")

                    For i As Integer = 0 To TrackNumAll.Length - 1

                        'TrackNum,TrxPartner
                        Dim TrackNumSplit As String() = TrackNumAll(i).Split(",")

                        Dim TrackNum As String = ""
                        Try
                            TrackNum = TrackNumSplit(0).ToUpper
                        Catch ex As Exception
                            TrackNum = ""
                        End Try
                        Dim TrxPartner As String = ""
                        Try
                            TrxPartner = TrackNumSplit(1).ToUpper
                        Catch ex As Exception
                            TrxPartner = ""
                        End Try

                        If TrackNum = "" Or TrxPartner = "" Then
                            Result(1) = "Input AWB dan TrxPartner"
                            GoTo Skip
                        End If

                        Query = "Select count(TrackNum)"
                        Query &= " From `TransactionDeliveryInfo`"
                        Query &= " Where TrackNum = '" & TrackNum & "'"
                        MCom.CommandText = Query
                        If MCom.ExecuteScalar < 1 Then
                            Result(1) = TrackNum & " tidak ditemukan"
                            GoTo Skip
                        End If

                    Next

                    'dari Case "CANCEL"


                Case "KEEP" 'penitipan kembali paket oleh DeliveryMan

                    'Sender = DeliveryMan
                    'Receiver = Toko

                    If ("" & ReceiverCode) = "" Then
                        Result(1) = "Input Kode Toko"
                        GoTo Skip
                    End If

                    If ValidateStoreCheckTransactionDeliveryInfo Then
                        If GetStore(MCon, ReceiverCode) = "" Then
                            Result(1) = GetStoreNotFound(MCon, ReceiverCode)
                            GoTo Skip
                        End If
                    End If

                    If SenderName = "" Then
                        Result(1) = "Input Nama Pengirim"
                        GoTo Skip
                    End If

                    Dim TrackNumAll As String() = TrackNumList.Split("|")

                    For i As Integer = 0 To TrackNumAll.Length - 1

                        'TrackNum,TrxPartner
                        Dim TrackNumSplit As String() = TrackNumAll(i).Split(",")

                        Dim TrackNum As String = ""
                        Try
                            TrackNum = TrackNumSplit(0).ToUpper
                        Catch ex As Exception
                            TrackNum = ""
                        End Try
                        Dim TrxPartner As String = ""
                        Try
                            TrxPartner = TrackNumSplit(1).ToUpper
                        Catch ex As Exception
                            TrxPartner = ""
                        End Try

                        If TrackNum = "" Or TrxPartner = "" Then
                            Result(1) = "Input AWB dan TrxPartner"
                            GoTo Skip
                        End If

                        Query = "Select count(TrackNum)"
                        Query &= " From `TransactionDeliveryInfo`"
                        Query &= " Where TrackNum = '" & TrackNum & "'"
                        MCom.CommandText = Query
                        If MCom.ExecuteScalar < 1 Then
                            Result(1) = TrackNum & " tidak ditemukan"
                            GoTo Skip
                        End If

                    Next

                    'dari Case "KEEP"


                Case Else

                    Result(1) = TrackStatus & " tidak dikenali"
                    GoTo Skip

            End Select 'dari Select Case TrackStatus

            '=== validasi end


            '=== begin proses sql

            Dim MTrn As MySqlTransaction

            MTrn = MCon.BeginTransaction()
            MCom.Transaction = MTrn

            Dim DebugCommitTransaction As Boolean = False
            If DebugCommitTransaction Then
                MTrn.Commit()
            End If


            'untuk mengatasi deadlock/timeout
            Dim MaxTry As Short = 1
            Dim nTry As Short = 0

            Dim MDa As New MySqlDataAdapter

            Select Case TrackStatus

                Case "PICKUP", "PICKUP_CANCEL" 'paket diambil di Toko oleh DeliveryMan

                    'Sender = Toko
                    'Receiver = DeliveryMan

                    Dim IsPickupCancel As Boolean = False
                    If TrackStatus.Contains("_CANCEL") Then
                        IsPickupCancel = True
                    End If

                    Dim AccountNeedTrackingBKO As String = CustomAccountKlikApka & "," & CustomAccountKlikIStore & "," & CustomAccountKlikFood & "," & CustomAccountKlikReOrder

                    Dim TrackNumAll As String() = TrackNumList.Split("|")

                    TrackNumList = ""

                    For i As Integer = 0 To TrackNumAll.Length - 1

                        'TrackNum,TrxPartner
                        Dim TrackNumSplit As String() = TrackNumAll(i).Split(",")

                        Dim TrackNum As String = TrackNumSplit(0).ToUpper
                        Dim TrxPartner As String = TrackNumSplit(0).ToUpper

                        If TrackNumList <> "" Then
                            TrackNumList &= "|"
                        End If
                        TrackNumList &= TrackNum

                        Query = "Update `transactiondeliveryinfo`"
                        If IsPickupCancel = False Then
                            Query &= " Set PINPickup = left(concat('X',PINPickup,'Y'),8)"
                        Else
                            Query &= " Set AddInfo = replace(concat(AddInfo,';PICKUPCANCEL=Y;'), ';;',';')"
                        End If
                        Query &= " , UpdTime = now(), UpdUser = 'Pickup_" & SenderCode & "'"
                        Query &= " Where TrackNum = '" & TrackNum & "'"
                        Try
                            MCom.CommandText = Query
                            MCom.ExecuteNonQuery()
                            TransactionDeliveryInfoTracking(AppName, AppVersion, Param, "PINPickup")
                        Catch ex As Exception
                            Dim e As New ClsError
                            e.ErrorLog(MCon, AppName, AppVersion, Method & " " & TrackStatus, SenderCode, ex, Query, LogKeyword)

                            Result(1) &= "Error : " & ex.Message
                            GoTo Skip
                        End Try


                        '== insert juga ke TrackingBKO BEGIN
                        Dim NeedTrackingBKO As Boolean = False

                        Query = "Select count(t.TrackNum) From `Transaction` t"
                        Query &= " Where t.TrackNum = @TrackNum"
                        Query &= " And t.ShAccount in ( " & AccountNeedTrackingBKO & " )"

                        MCom.Parameters.Clear()
                        MCom.Parameters.AddWithValue("@TrackNum", TrackNum)

                        Try
                            MCom.CommandText = Query
                            If MCom.ExecuteScalar > 0 Then
                                NeedTrackingBKO = True
                            End If
                        Catch ex As Exception
                            NeedTrackingBKO = False
                        End Try


                        If NeedTrackingBKO Then

                            'sesuai janjian dengan it toko, yang dicatat adalah yang terjadi pada hari terkecil, jam maximal
                            'contoh proses PIU atas awb yang sama
                            '1. 2021-09-06 09:00:00 -> dicatat pertama kali, tapi kemudian di-overwrite oleh nomor 2
                            '2. 2021-09-06 15:30:00 -> yang akan dicatat dan dilaporkan ke backoffice
                            '3. 2021-09-07 12:00:00 -> dicatat, tapi TrackNum di-edit supaya tidak terbaca ke backoffice

                            Query = "Select concat( MinTrackTime,'|'"
                            Query &= " , cast(ifnull(curdate() <= date(MinTrackTime),'-1') as char) )"
                            Query &= " From ("
                            Query &= " Select cast(ifnull(min(TrackTime),'') as char) as MinTrackTime"
                            Query &= " From TrackingBKO Where TrackNum = @TrackNum And `Status` = 'PIU'"
                            Query &= " )x"
                            MCom.CommandText = Query

                            MCom.Parameters.Clear()
                            MCom.Parameters.AddWithValue("@TrackNum", TrackNum)

                            Dim PIUTrackTimeInfo As String = ""
                            Try
                                PIUTrackTimeInfo = ("" & MCom.ExecuteScalar).ToString.Trim
                            Catch ex As Exception
                                PIUTrackTimeInfo = ""
                            End Try

                            If PIUTrackTimeInfo = "" Then
                                PIUTrackTimeInfo = "|-1"
                            End If

                            Dim MinTrackTimePIU As String = ""
                            Try
                                MinTrackTimePIU = PIUTrackTimeInfo.Split("|")(0)
                            Catch ex As Exception
                                MinTrackTimePIU = ""
                            End Try

                            If MinTrackTimePIU <> "" Then

                                Dim TrackTimeSameDate As String = "0"
                                Try
                                    TrackTimeSameDate = PIUTrackTimeInfo.Split("|")(1)
                                Catch ex As Exception
                                    TrackTimeSameDate = "0"
                                End Try

                                MCom.Parameters.Clear()

                                If TrackTimeSameDate = "1" Then
                                    'PIU di tanggal yang sama

                                    'Update informasi TrackingBKO saja
                                    Query = "Update TrackingBKO"
                                    Query &= " Set TrxPartner = @TrxPartner, TrackTime = now()"

                                    'refresh juga kolom yang lain - 230919
                                    Query &= " , AddInfo = @AddInfo"
                                    Query &= " , Company1 = @SenderCompany, Opr1 = @SenderType, OprId1 = @SenderID, OprName1 = @SenderName"
                                    Query &= " , Company2 = @ReceiverCompany, Opr2 = @ReceiverType, OprId2 = @ReceiverID, OprName2 = @ReceiverName"
                                    Query &= " , TrackUser = @SenderType, TrackUserID = @SenderCode"

                                    Query &= " Where TrackNum = @TrackNum And `Status` = 'PIU'"

                                    MCom.Parameters.AddWithValue("@TrackNum", TrackNum)
                                    MCom.Parameters.AddWithValue("@TrxPartner", TrxPartner)

                                    'refresh juga kolom yang lain - 230919
                                    MCom.Parameters.AddWithValue("@AddInfo", TrackDate & " " & TrackTime)
                                    MCom.Parameters.AddWithValue("@SenderCompany", SenderCompany)
                                    MCom.Parameters.AddWithValue("@SenderType", SenderType)
                                    MCom.Parameters.AddWithValue("@SenderID", SenderID)
                                    MCom.Parameters.AddWithValue("@SenderName", SenderName)
                                    MCom.Parameters.AddWithValue("@ReceiverCompany", ReceiverCompany & "|" & IIf(IsPickupCancel, TrackStatus, ""))
                                    MCom.Parameters.AddWithValue("@ReceiverType", ReceiverType)
                                    MCom.Parameters.AddWithValue("@ReceiverID", ReceiverID)
                                    MCom.Parameters.AddWithValue("@ReceiverName", ReceiverName)
                                    MCom.Parameters.AddWithValue("@SenderCode", SenderCode)

                                Else
                                    'PIU di tanggal yang sudah lebih maju

                                    'Insert Tracking BKO PIU
                                    Query = "Insert Into TrackingBKO ("
                                    Query &= " TrackNum, `Status`, ForCustomer"
                                    Query &= " , ForCustomerIna, ForCustomerEng"
                                    Query &= " , AddInfo, TrxPartner"
                                    Query &= " , Company1, Opr1, OprId1, OprName1"
                                    Query &= " , Company2, Opr2, OprId2, OprName2"
                                    Query &= " , TrackTime, TrackUser, TrackUserID"
                                    Query &= " ) Select"
                                    Query &= " @TrackNum, 'PIU', '1'"
                                    Query &= " , 'PIN AMBIL', 'PICKUP PIN'"
                                    Query &= " , @AddInfo, @TrxPartner"
                                    Query &= " , @SenderCompany, @SenderType, @SenderID, @SenderName"
                                    Query &= " , @ReceiverCompany, @ReceiverType, @ReceiverID, @ReceiverName"
                                    Query &= " , now(), @SenderType, @SenderCode"

                                    MCom.Parameters.AddWithValue("@TrackNum", TrackNum & "X")
                                    MCom.Parameters.AddWithValue("@TrxPartner", TrxPartner)
                                    MCom.Parameters.AddWithValue("@AddInfo", TrackDate & " " & TrackTime)
                                    MCom.Parameters.AddWithValue("@SenderCompany", SenderCompany)
                                    MCom.Parameters.AddWithValue("@SenderType", SenderType)
                                    MCom.Parameters.AddWithValue("@SenderID", SenderID)
                                    MCom.Parameters.AddWithValue("@SenderName", SenderName)
                                    MCom.Parameters.AddWithValue("@ReceiverCompany", ReceiverCompany & "|" & IIf(IsPickupCancel, TrackStatus, ""))
                                    MCom.Parameters.AddWithValue("@ReceiverType", ReceiverType)
                                    MCom.Parameters.AddWithValue("@ReceiverID", ReceiverID)
                                    MCom.Parameters.AddWithValue("@ReceiverName", ReceiverName)
                                    MCom.Parameters.AddWithValue("@SenderCode", SenderCode)

                                End If 'dari If TrackTimeSameDate = "1"

                                Try
                                    MCom.CommandText = Query
                                    MCom.ExecuteNonQuery()
                                Catch ex As Exception
                                    Dim e As New ClsError
                                    e.ErrorLog(MCon, AppName, AppVersion, Method & " " & TrackStatus, SenderCode, ex, Query, LogKeyword)

                                    Result(1) &= "Error : " & ex.Message
                                    GoTo Skip
                                End Try

                            Else

                                'Insert Tracking BKO PIU
                                Query = "Insert Into TrackingBKO ("
                                Query &= " TrackNum, `Status`, ForCustomer"
                                Query &= " , ForCustomerIna, ForCustomerEng"
                                Query &= " , AddInfo, TrxPartner"
                                Query &= " , Company1, Opr1, OprId1, OprName1"
                                Query &= " , Company2, Opr2, OprId2, OprName2"
                                Query &= " , TrackTime, TrackUser, TrackUserID"
                                Query &= " ) Select"
                                Query &= " @TrackNum, 'PIU', '1'"
                                Query &= " , 'PIN AMBIL', 'PICKUP PIN'"
                                Query &= " , @AddInfo, @TrxPartner"
                                Query &= " , @SenderCompany, @SenderType, @SenderID, @SenderName"
                                Query &= " , @ReceiverCompany, @ReceiverType, @ReceiverID, @ReceiverName"
                                Query &= " , now(), @SenderType, @SenderCode"

                                MCom.Parameters.Clear()
                                MCom.Parameters.AddWithValue("@TrackNum", TrackNum)
                                MCom.Parameters.AddWithValue("@TrxPartner", TrxPartner)
                                MCom.Parameters.AddWithValue("@AddInfo", TrackDate & " " & TrackTime)
                                MCom.Parameters.AddWithValue("@SenderCompany", SenderCompany)
                                MCom.Parameters.AddWithValue("@SenderType", SenderType)
                                MCom.Parameters.AddWithValue("@SenderID", SenderID)
                                MCom.Parameters.AddWithValue("@SenderName", SenderName)
                                MCom.Parameters.AddWithValue("@ReceiverCompany", ReceiverCompany & "|" & IIf(IsPickupCancel, TrackStatus, ""))
                                MCom.Parameters.AddWithValue("@ReceiverType", ReceiverType)
                                MCom.Parameters.AddWithValue("@ReceiverID", ReceiverID)
                                MCom.Parameters.AddWithValue("@ReceiverName", ReceiverName)
                                MCom.Parameters.AddWithValue("@SenderCode", SenderCode)

                                Try
                                    MCom.CommandText = Query
                                    MCom.ExecuteNonQuery()
                                Catch ex As Exception
                                    Dim e As New ClsError
                                    e.ErrorLog(MCon, AppName, AppVersion, Method & " " & TrackStatus, SenderCode, ex, Query, LogKeyword)

                                    Result(1) &= "Error : " & ex.Message
                                    GoTo Skip
                                End Try

                            End If 'dari If MinTrackTimePIU <> ""

                        End If 'dari If NeedTrackingBKO
                        '== insert juga ke TrackingBKO END

                        If IsPickupCancel = True Then
                            'Insert Push Tracking DMS PUC
                            Query = "Insert Into PushTracking ("
                            Query &= " Account, TrackNum, TrackStatus"
                            Query &= " , TrackTime, TrackUser, TrackUserID"
                            Query &= " , Company1, Opr1, OprID1, OprName1"
                            Query &= " , AddInfo"
                            Query &= " , PushStatus, AddTime, AddUser"
                            Query &= " ) values ("
                            Query &= " 'DMS', @TrackNum, 'PUC'"
                            Query &= " , now(), @SenderType, @SenderCode"
                            Query &= " , @SenderCompany, @SenderType, @SenderID, @SenderName"
                            Query &= " , @TrxPartner"
                            Query &= " , 0, now(), @Method"
                            Query &= " )"
                            Try
                                MCom.CommandText = Query

                                MCom.Parameters.Clear()
                                MCom.Parameters.AddWithValue("@TrackNum", TrackNum)
                                MCom.Parameters.AddWithValue("@SenderCompany", SenderCompany)
                                MCom.Parameters.AddWithValue("@SenderType", SenderType)
                                MCom.Parameters.AddWithValue("@SenderID", SenderID)
                                MCom.Parameters.AddWithValue("@SenderName", SenderName)
                                MCom.Parameters.AddWithValue("@SenderCode", SenderCode)
                                MCom.Parameters.AddWithValue("@TrxPartner", TrxPartner)
                                MCom.Parameters.AddWithValue("@Method", Left(Method, 30))

                                MCom.ExecuteNonQuery()
                            Catch ex As Exception
                                Dim e As New ClsError
                                e.ErrorLog(MCon, AppName, AppVersion, Method & " " & TrackStatus, SenderCode, ex, Query, LogKeyword)

                                Result(1) &= "Error : " & ex.Message
                                GoTo Skip
                            End Try

                        End If 'dari If IsPickupCancel = True

                    Next

                    Dim TrackNumListSQL As String = ConvertNumListToSQL(TrackNumList)

                    'dari Case "PICKUP"


                Case "RETURN" 'pengembalian bulky ke Toko oleh DeliveryMan

                    'Sender = DeliveryMan
                    'Receiver = Toko

                    Dim TrackNumAll As String() = TrackNumList.Split("|")

                    TrackNumList = ""

                    For i As Integer = 0 To TrackNumAll.Length - 1

                        'TrackNum,TrxPartner
                        Dim TrackNumSplit As String() = TrackNumAll(i).Split(",")

                        Dim TrackNum As String = TrackNumSplit(0).ToUpper
                        Dim TrxPartner As String = TrackNumSplit(1).ToUpper

                        If TrackNumList <> "" Then
                            TrackNumList &= "|"
                        End If
                        TrackNumList &= TrackNum

                        Query = "Update `transactiondeliveryinfo`"
                        Query &= " Set PINReturn = left(concat('X',PINReturn,'Y'),8)"
                        Query &= " , BasePin = null"
                        Query &= " , PINPickup = '', PINCancel = '', PINKeep = ''"
                        Query &= " , UpdTime = now(), UpdUser = 'Return_" & ReceiverCode & "'"
                        Query &= " Where TrackNum = '" & TrackNum & "'"
                        Try
                            MCom.CommandText = Query
                            MCom.ExecuteNonQuery()
                            TransactionDeliveryInfoTracking(AppName, AppVersion, Param, "PINReturn")
                        Catch ex As Exception
                            Dim e As New ClsError
                            e.ErrorLog(MCon, AppName, AppVersion, Method & " " & TrackStatus, ReceiverCode, ex, Query, LogKeyword)

                            Result(1) &= "Error : " & ex.Message
                            GoTo Skip
                        End Try

                    Next

                    'dari Case "RETURN"


                Case "CANCEL" 'pembatalan pengiriman oleh DeliveryMan

                    'Sender = DeliveryMan
                    'Receiver = Toko

                    Dim TrackNumAll As String() = TrackNumList.Split("|")

                    TrackNumList = ""

                    For i As Integer = 0 To TrackNumAll.Length - 1

                        'TrackNum,TrxPartner
                        Dim TrackNumSplit As String() = TrackNumAll(i).Split(",")

                        Dim TrackNum As String = TrackNumSplit(0).ToUpper
                        Dim TrxPartner As String = TrackNumSplit(1).ToUpper

                        If TrackNumList <> "" Then
                            TrackNumList &= "|"
                        End If
                        TrackNumList &= TrackNum

                        Query = "Update `transactiondeliveryinfo`"
                        Query &= " Set PINCancel = left(concat('X',PINCancel,'Y'),8)"
                        Query &= " , BasePin = null"
                        Query &= " , PINPickup = '', PINReturn = '', PINKeep = ''"
                        Query &= " , UpdTime = now(), UpdUser = 'Cancel_" & ReceiverCode & "'"
                        Query &= " Where TrackNum = '" & TrackNum & "'"
                        Try
                            MCom.CommandText = Query
                            MCom.ExecuteNonQuery()
                            TransactionDeliveryInfoTracking(AppName, AppVersion, Param, "PINCancel")
                        Catch ex As Exception
                            Dim e As New ClsError
                            e.ErrorLog(MCon, AppName, AppVersion, Method & " " & TrackStatus, ReceiverCode, ex, Query, LogKeyword)

                            Result(1) &= "Error : " & ex.Message
                            GoTo Skip
                        End Try

                    Next

                    'dari Case "CANCEL"


                Case "KEEP" 'penitipan kembali paket oleh DeliveryMan

                    'Sender = DeliveryMan
                    'Receiver = Toko

                    Dim TrackNumAll As String() = TrackNumList.Split("|")

                    TrackNumList = ""

                    For i As Integer = 0 To TrackNumAll.Length - 1

                        'TrackNum,TrxPartner
                        Dim TrackNumSplit As String() = TrackNumAll(i).Split(",")

                        Dim TrackNum As String = TrackNumSplit(0).ToUpper
                        Dim TrxPartner As String = TrackNumSplit(1).ToUpper

                        If TrackNumList <> "" Then
                            TrackNumList &= "|"
                        End If
                        TrackNumList &= TrackNum

                        Query = "Update `transactiondeliveryinfo`"
                        Query &= " Set PINKeep = left(concat('X',PINKeep,'Y'),8)"
                        'Query &= " , BasePin = null"'ada kemungkinan akan dipakai lagi untuk manual order
                        Query &= " , PINReturn = '', PINCancel = ''"
                        Query &= " , UpdTime = now(), UpdUser = 'Keep_" & ReceiverCode & "'"
                        Query &= " Where TrackNum = '" & TrackNum & "'"
                        Try
                            MCom.CommandText = Query
                            MCom.ExecuteNonQuery()
                            TransactionDeliveryInfoTracking(AppName, AppVersion, Param, "PINKeep")
                        Catch ex As Exception
                            Dim e As New ClsError
                            e.ErrorLog(MCon, AppName, AppVersion, Method & " " & TrackStatus, ReceiverCode, ex, Query, LogKeyword)

                            Result(1) &= "Error : " & ex.Message
                            GoTo Skip
                        End Try

                    Next

                    'dari Case "KEEP"


                Case Else

                    Result(1) = TrackStatus & " tidak dikenali"
                    GoTo Skip

            End Select 'dari Select Case TrackStatus


            '=== proses sql end


            Process = True

Skip:

            If DebugCommitTransaction Then
                MTrn = MCon.BeginTransaction()
                MCom.Transaction = MTrn
            End If

            If Process Then

                Result(0) = "0"
                Result(1) = ResponseOK
                Result(2) = ""

                Try
                    MTrn.Commit()
                Catch ex As Exception
                End Try

            Else

                Try
                    MTrn.Rollback()
                Catch ex As Exception
                End Try

            End If

            CreateResponseLog(AppName, AppVersion, Method & " " & TrackStatus, User, Result, , LogKeyword)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method & " " & TrackStatus, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Public Function TransactionDeliveryInfoTracking(ByVal AppName As String, ByVal AppVersion As String _
    , ByVal Param As Object(), ByVal Type As String)
        Dim Mcon As New MySqlConnection
        Dim Method As String = "TransactionDeliveryInfoTracking"

        Dim Query As String = ""
        Dim Result As Object() = CreateResult()

        Dim User As String = Param(0).ToString
        Dim Password As String = Param(1).ToString

        Dim service As New LocalCore

        Try
            Mcon = MasterMCon.Clone
            'Dim UserOK As Object() = service.ValidasiUser(Mcon, User, Password)

            'If UserOK(0) <> "0" Then
            '    Result(1) = UserOK(1)
            '    GoTo Skip
            'End If

            SqlParam = New Dictionary(Of String, String)

            Dim ObjSQL As New ClsSQL
            Dim Value As String = ""

            SqlParam.Add("@TrackNumList", "" & Param(14))
            SqlParam.Add("@TrackStatus", "" & Param(3))
            SqlParam.Add("@TrackDateTime", "" & Param(2))
            SqlParam.Add("@PinType", "" & Type)
            SqlParam.Add("@SenderCompany", "" & Param(4))
            SqlParam.Add("@SenderName", "" & Param(8))
            SqlParam.Add("@SenderCode", "" & Param(6))
            SqlParam.Add("@SenderType", "" & Param(5))
            SqlParam.Add("@SenderID", "" & Param(7))
            SqlParam.Add("@ReceiverCompany", "" & Param(9))
            SqlParam.Add("@ReceiverName", "" & Param(13))
            SqlParam.Add("@ReceiverCode", "" & Param(11))
            SqlParam.Add("@ReceiverType", "" & Param(10))

            Query = "SELECT COUNT(TrackNumList) as SeqNo From TransactionDeliveryInfoTracking"
            Query &= " WHERE TrackNumList = @TrackNumList"
            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(Mcon, Query, SqlParam)
            SqlParam.Add("@SeqNo", "" & dtQuery.Rows(0).Item("SeqNo") + 1)

            If Type = "PINPickup" Then
                Query = "SELECT PINPickup From transactiondeliveryinfo Where TrackNum = @TrackNumList"
                dtQuery = ObjSQL.ExecDatatableWithParam(Mcon, Query, SqlParam)
                SqlParam.Add("@Pin", "" & dtQuery.Rows(0).Item("PINPickup").ToString)

            ElseIf Type = "PINReturn" Then
                Query = "SELECT PINReturn From transactiondeliveryinfo Where TrackNum = @TrackNumList"
                dtQuery = ObjSQL.ExecDatatableWithParam(Mcon, Query, SqlParam)
                SqlParam.Add("@Pin", "" & dtQuery.Rows(0).Item("PINReturn").ToString)

            ElseIf Type = "PINCancel" Then
                Query = "SELECT PINCancel From transactiondeliveryinfo Where TrackNum = @TrackNumList"
                dtQuery = ObjSQL.ExecDatatableWithParam(Mcon, Query, SqlParam)
                SqlParam.Add("@Pin", "" & dtQuery.Rows(0).Item("PINCancel").ToString)

            ElseIf Type = "PINKeep" Then
                Query = "SELECT PINKeep From transactiondeliveryinfo Where TrackNum = @TrackNumList"
                dtQuery = ObjSQL.ExecDatatableWithParam(Mcon, Query, SqlParam)
                SqlParam.Add("@Pin", "" & dtQuery.Rows(0).Item("PINPickup").ToString)
            End If

            Query = ""
            Query = "INSERT INTO TransactionDeliveryInfoTracking("
            Query &= "TrackNumList, TrackStatus, TrackDateTime, SeqNo, Pin, PinType, SenderCompany, SenderName"
            Query &= ", SenderCode, SenderType, SenderID, ReceiverCompany, ReceiverName, ReceiverCode"
            Query &= ", ReceiverType) VALUES (@TrackNumList, @TrackStatus, @TrackDateTime, @SeqNo, @Pin"
            Query &= ", @PinType, @SenderCompany, @SenderName, @SenderCode, @SenderType, @SenderID"
            Query &= ", @ReceiverCompany, @ReceiverName, @ReceiverCode, @ReceiverType)"

            If ObjSQL.ExecNonQueryWithParam(Mcon, Query, SqlParam) Then
                Result(0) = 0
                Result(1) = "Berhasil"
                Result(2) = ""
            Else
                Result(1) = "Gagal Query"
            End If

Skip:
        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)
            Result(1) &= "Error : " & ex.Message
        Finally
            If Mcon.State <> ConnectionState.Closed Then
                Mcon.Close()
            End If
            Try
                Mcon.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result
    End Function

    Public Function GetIndomaretDCDepoList(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            MCon = MasterMCon.Clone

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            '    Result(1) = UserOK(1)
            '    GoTo Skip
            'End If

            Dim ObjSQL As New ClsSQL

            Query = "Select `Value`, Display, `Type` From ("

            Query &= " Select i.`Code` as `Value`"
            Query &= " , concat(case when ifnull(i.Alias,'') <> '' then i.Alias else i.Name end,' (',i.Code,')') as Display"
            Query &= " , i.Coverage, 'DCI' as `Type`"
            Query &= " From indomaretdc i"
            Query &= " Where ucase(`Type`) = 'IDM'"
            Query &= " And curdate() between i.ActiveDate and ifnull(i.InActiveDate,curdate())"

            Query &= " union"

            Query &= " Select i.`Code` as `Value`"
            Query &= " , concat(case when ifnull(i.Alias,'') <> '' then i.Alias else i.Name end,' (',i.Code,')') as Display"
            Query &= " , i.Coverage, 'DPO' as `Type`"
            Query &= " From indomaretdepo i"
            Query &= " Where 1=1"
            Query &= " And curdate() between i.ActiveDate and ifnull(i.InActiveDate,curdate())"

            Query &= " )x"
            Query &= " Order by x.Coverage, x.Display"

            Dim dtQuery As DataTable = ObjSQL.ExecDatatable(MCon, Query)

            If dtQuery Is Nothing Then
                Result(1) = "Gagal query"
                GoTo Skip
            End If

            If dtQuery.Rows.Count < 1 Then
                'Result(1) = "Tidak ditemukan"
                'GoTo Skip
            End If

            If Not dtQuery Is Nothing Then

                Result(2) = ConvertDatatableToString(dtQuery)
                Result(1) = ResponseOK
                Result(0) = "0"

            End If

Skip:

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetIndomaretDCDepoList", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Public Function GetRackInStoreList(ByVal AppName As String, ByVal AppVersion As String _
        , ByVal DCCode As String, ByVal Param As Object()) As Object()

        DCCode = DCCode.ToUpper

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString
            Dim StoreList As String = Param(2).ToString.Trim
            Dim TrackNumList As String = Param(3).ToString.Trim

            MCon = MasterMCon.Clone

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'GoTo Skip
            'End If

            If DCCode = "" Then
                Result(1) = "Input Kode DC"
                GoTo Skip
            End If

            Dim Hub As String = GetMappingHubDC(MCon, DCCode)
            If Hub = "" Then
                Result(1) = GetMappingHubDCNotFound(DCCode)
                GoTo Skip
            End If

            If StoreList = "" Then
                Result(1) = "Input daftar Toko"
                GoTo Skip
            End If

            StoreList = ConvertNumListToSQL(StoreList.Replace(",", "|"))


            Dim ObjSQL As New ClsSQL

            Dim IncludeDcTempRequest As Boolean = False

            Query = "Select r.TrackNum"
            Query &= " , r.Dst as Store, r.Rack as Zone"
            Query &= " From RackIn r"
            Query &= " Where r.Hub = '" & Hub & "'"
            If TrackNumList <> "" Then
                TrackNumList = ConvertNumListToSQL(TrackNumList.Replace(",", "|"))
                Query &= " AND r.TrackNum in (" & TrackNumList & ")"
            End If
            Query &= " And r.Dst in (" & StoreList & ")"
            Query &= " And r.TrackNum not in"
            Query &= " (Select TrackNum From raodraft"
            Query &= " Where Hub = '" & Hub & "')"
            If IncludeDcTempRequest Then
                Query &= " And r.TrackNum not in"
                Query &= " (Select TrackNum From DCTempRequest"
                Query &= " Where Hub = '" & Hub & "')"
            End If
            Query &= " Order By length(r.Rack), r.Rack, r.Dst, r.TrackNum"

            Dim dtQuery As DataTable = ObjSQL.ExecDatatable(MCon, Query)

            If dtQuery Is Nothing Then
                Result(1) = "Gagal query"
                GoTo Skip
            End If

            Result(2) = ConvertDatatableToString(dtQuery)
            Result(1) = ResponseOK
            Result(0) = "0"

Skip:

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetRackInStoreList", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Private Function GetMappingHubDC(ByVal MCon As MySqlConnection, ByVal DCCode As String, Optional ByVal IncludeIGR As Boolean = False) As String

        Dim Hub As String = ""

        Try
            Dim ObjSQL As New ClsSQL
            'Hub = "" & ObjSQL.ExecScalar(MCon, "Select Hub From MappingHubDC Where DC = '" & DCCode & "'" & IIf(IncludeIGR, " or IGR = '" & DCCode & "'", ""))

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@DCCode", DCCode)

            Hub = "" & ObjSQL.ExecScalarWithParam(MCon, "Select Hub From MappingHubDC Where DC = @DCCode" & IIf(IncludeIGR, " or IGR = @DCCode", ""), SqlParam)

        Catch ex As Exception
            Hub = ""

        End Try

        Return Hub

    End Function

    Private Function GetMappingHubDCNotFound(ByVal DCCode As String) As String
        Return "Hub untuk DC " & DCCode & " tidak ditemukan"
    End Function

    Public Function DCPickingManual(ByVal AppName As String, ByVal AppVersion As String _
    , ByVal DCCode As String, ByVal Param As Object()) As Object()

        DCCode = DCCode.ToUpper

        Dim Method As String = "DCPickingManual"

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString
            Dim StoreList As String = Param(2).ToString.Trim.ToUpper

            Dim CustomType As String = ""
            Try
                CustomType = Param(3).ToString.Trim.ToUpper
            Catch ex As Exception
                CustomType = ""
            End Try

            If CustomType = "" Then
                CustomType = "ROTI"
            End If


            Dim CustomPrefix As String = ""
            If CustomType = "KURWIL" Then
                CustomPrefix = "K"
            End If
            If CustomPrefix = "" Then
                CustomPrefix = "R"
            End If


            Dim CustomRefNo As String = ""
            Try
                CustomRefNo = Param(4).ToString.Trim.ToUpper
            Catch ex As Exception
                CustomRefNo = ""
            End Try

            Dim IsManual As String = ""
            Try
                IsManual = Param(5).ToString
                If IsNumeric(IsManual) = False Then
                    IsManual = ""
                End If
            Catch ex As Exception
                IsManual = ""
            End Try
            If IsManual = "" Then
                IsManual = "1" 'default adalah proses manual dari webreport
            End If

            Dim TrackNumList As String = Param(6).ToString.Trim.ToUpper

            MCon = MasterMCon.Clone
            MCon.Open()

            Dim LogKeyword As String = DCCode & " " & StoreList
            LogKeyword = Strings.Left(LogKeyword, 100)

            CreateRequestLog(AppName, AppVersion, Method, User, DCCode, Param, LogKeyword)

            Dim Hub As String = GetMappingHubDC(MCon, DCCode)

            Dim ObjSQL As New ClsSQL

            'jagaan validasi zona toko
            Dim ValidasiZonaToko As Boolean = (CustomType <> "KURWIL")
            If ValidasiZonaToko Then
                'StoreList = ConvertNumListToSQL(StoreList.Replace(",", "|"))

                Query = "Select count(distinct(r.Rack))"
                Query &= " From RackIn r"
                Query &= " Where r.Hub = @Hub"
                Query &= " And r.Dst in (" & ConvertNumListToSQL(StoreList.Replace(",", "|")) & ")"
                If TrackNumList <> "" Then
                    Query &= " And r.TrackNum in (" & ConvertNumListToSQL(TrackNumList.Replace(",", "|")) & ")"
                End If
                Query &= " And r.TrackNum not in"
                Query &= " (Select TrackNum From raodraft"
                Query &= " Where Hub = @Hub)"
                Query &= " Order By length(r.Rack), r.Rack, r.Dst, r.TrackNum"

                SqlParam = New Dictionary(Of String, String)
                SqlParam.Add("@Hub", Hub)

                If ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam) > 1 Then
                    Result(1) = "Pastikan toko ada pada zona yang sama!"
                    GoTo Skip
                End If
            End If 'dari ValidasiZonaToko

            'sudah pernah ada picking draft sebelumnya
            Query = "Select TrackNum From DCTempRequest "
            Query &= " Where Hub = @Hub"
            If TrackNumList <> "" Then
                Query &= " And TrackNum in (" & ConvertNumListToSQL(TrackNumList.Replace(",", "|")) & ")"
            End If
            Query &= " And instr('" & StoreList.Replace("|", ",") & "',Dst)" 'daftar toko yang di-request

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Hub", Hub)

            Dim dtDraft As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)

            If Not dtDraft Is Nothing Then

                Dim sb As New StringBuilder

                For d As Integer = 0 To dtDraft.Rows.Count - 1
                    If d <> 0 Then
                        sb.Append("|")
                    End If
                    sb.Append(dtDraft.Rows(d).Item("TrackNum").ToString)
                Next
                Dim TrackNumDraft As String = sb.ToString

                If TrackNumDraft <> "" Then
                    Query = "Delete From DCTempRequest "
                    Query &= " Where Hub = @Hub"
                    Query &= " And instr('" & StoreList.Replace("|", ",") & "',Dst)" 'daftar toko yang di-request

                    SqlParam = New Dictionary(Of String, String)
                    SqlParam.Add("@Hub", Hub)

                    ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

                    'Dim e As New ClsError
                    'e.DebugLog(AppName, AppVersion, Method, User, "Hapus Picking Draft " & TrackNumDraft.Replace("|", ","))
                End If 'dari If TrackNumDraft <> ""

            End If 'dari If Not dtDraft Is Nothing


            Dim dParam(4) As Object
            dParam(0) = User
            dParam(1) = Password
            dParam(2) = StoreList
            dParam(3) = CustomType
            dParam(4) = TrackNumList

            Result = DCPickingDraft(AppName, AppVersion, DCCode, dParam)

            Dim DateNow As DateTime = Date.Now

            Dim PickupNo As String = CustomPrefix & Strings.Right(DCCode, 3) & Format(DateNow, "yyMMddHHmmss")

            Dim fParam(9) As Object
            fParam(0) = User
            fParam(1) = Password
            fParam(2) = PickupNo
            fParam(3) = DateNow
            fParam(4) = StoreList
            fParam(5) = CustomType
            'fParam(6)
            fParam(7) = CustomRefNo
            fParam(8) = IsManual
            fParam(9) = TrackNumList

            Result = DCPickingFinal(AppName, AppVersion, DCCode, fParam)

            If Result(0).ToString = "0" Then

                StoreList = ConvertNumListToSQL(StoreList.Replace(",", "|"))

                Query = "Select Distinct PickupNo From RaoDraft"
                Query &= " Where Hub = @Hub and Dst in (" & StoreList & ")"
                Query &= " Order By Updtime Desc"
                Query &= " Limit 1"

                SqlParam = New Dictionary(Of String, String)
                SqlParam.Add("@Hub", Hub)

                Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)

                If dtQuery Is Nothing Then
                    Result(1) = "Gagal query PickupNo"
                    GoTo Skip
                End If

                Result(2) = ConvertDatatableToString(dtQuery)
                Result(1) = ResponseOK
                Result(0) = "0"

            End If

Skip:

            CreateResponseLog(AppName, AppVersion, Method, User, Result, PickupNo, LogKeyword)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Public Function DCPickingFinal(ByVal AppName As String, ByVal AppVersion As String _
    , ByVal DCCode As String, ByVal Param As Object()) As Object()

        DCCode = DCCode.ToUpper

        Dim Method As String = "DCPickingFinal"

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            Dim PickupNo As String = Param(2).ToString.Trim.ToUpper
            Dim PickupTime As DateTime = Param(3)
            Dim StoreList As String = Param(4).ToString.Trim.ToUpper
            Dim TrackNumList As String = Param(9).ToString.Trim.ToUpper

            Dim tTipe As String = ""
            Try
                tTipe = ("" & Param(5)).ToString.ToUpper.Trim
            Catch ex As Exception
                tTipe = ""
            End Try

            Dim StoreInfoList As String = ""
            Try
                StoreInfoList = ("" & Param(6)).ToString.ToUpper.Trim
            Catch ex As Exception
                StoreInfoList = ""
            End Try

            Dim CustomRefNo As String = ""
            Try
                CustomRefNo = ("" & Param(7)).ToString.ToUpper.Trim
            Catch ex As Exception
                CustomRefNo = ""
            End Try

            Dim IsManual As Boolean = False
            Try
                IsManual = (("" & Param(8)).ToString.ToUpper.Trim = "1")
            Catch ex As Exception
                IsManual = False
            End Try

            Dim Tipe As String = "0" 'default = DRY

            Dim MTrn As MySqlTransaction

            MCon = MasterMCon.Clone

            Dim MCom As New MySqlCommand("", MCon)
            MCon.Open()

            'Dim x As New ClsError
            'Dim tStart As DateTime = Date.Now()
            'Dim tEnd As DateTime = Date.Now()
            'Dim tInterval As Long = DateDiff(DateInterval.Second, tStart, tEnd)

            Dim LogKeyword As String = DCCode & " " & PickupNo

            CreateRequestLog(AppName, AppVersion, Method, User, DCCode, Param, LogKeyword)

            Dim Process As Boolean = False

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'GoTo Skip
            'End If

            Dim CheckDCCode As String = GetDC(MCon, DCCode)
            If CheckDCCode = "" Then
                CheckDCCode = GetDepo(MCon, DCCode)
            End If

            If CheckDCCode = "" Then
                Result(1) = GetDCDepoNotFound(DCCode)
                GoTo Skip
            End If

            Dim Hub As String = GetMappingHubDC(MCon, DCCode, True)
            If Hub = "" Then
                Result(1) = GetMappingHubDCNotFound(DCCode)
                GoTo Skip
            End If

            If PickupNo = "" Then
                Result(1) = "Input PickupNo"
                GoTo Skip
            End If

            If StoreList = "" Then
                Result(1) = "Input Daftar Toko"
                GoTo Skip
            End If

            'Dim Tipe As String = "0"
            'If PickupNo.ToUpper.StartsWith("R") Then
            '    Tipe = "1" 'karena picking roti juga pakai draft dan final
            'End If

            If tTipe <> "" Then
                If tTipe.Contains("ROTI") Then
                    Tipe = "1"
                End If
                If tTipe.Contains("KURWIL") Then
                    Tipe = "2"
                End If
            Else
                If PickupNo.ToUpper.StartsWith("R") Then
                    Tipe = "1"
                End If
                If PickupNo.ToUpper.StartsWith("K") Then
                    Tipe = "2"
                End If
            End If


            Query = "Select Count(rao.PickupNo) From RaoDraft rao, DCRequest dcr"
            Query &= " Where rao.Hub = @Hub And rao.PickupNo = @PickupNo"
            Query &= " And rao.PickupNo = dcr.PickupNo And rao.Hub = dcr.Hub And dcr.`Type` = @Tipe"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Hub", Hub)
            SqlParam.Add("@PickupNo", PickupNo)
            SqlParam.Add("@Tipe", Tipe)

            Dim ObjSQL As New ClsSQL
            If ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam) > 0 Then 'sudah pernah request picking, jadi tinggal balikin daftar toko dengan nomor resi
                Process = True
                GoTo Skip
            End If

            Dim DCIGR As String = ""
            Query = "Select DC From MappingHubDC Where IGR = @DCCode"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@DCCode", DCCode)

            Try
                DCIGR = ("" & ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam)).ToString.Trim.ToUpper
            Catch ex As Exception
                DCIGR = ""
            End Try

            If DCIGR = "" Then
                DCIGR = DCCode
            End If


            Dim PickingSource As String = ""
            If IsManual Or Tipe = "2" Then
                PickingSource = "IPP"
            ElseIf DCIGR <> "" And DCIGR <> DCCode Then
                PickingSource = "IGR"
            Else
                PickingSource = "IDM"
            End If

            Dim PickingType As String = ""
            If Tipe = "2" Then
                PickingType = "KUW"
            ElseIf Tipe = "1" Then
                PickingType = "ROT"
            Else
                PickingType = "DRY"
            End If

            Dim PickingManual As String = ""
            If IsManual = False Then
                PickingManual = "SYS"
            Else
                PickingManual = "MNL"
            End If

            Dim TrackingAddInfo As String = PickingSource & "|" & PickingType & "|" & PickingManual
            'Source = IDM/IGR/IPP
            'Type   = DRY/ROT/KUW
            'Manual = SYS/MNL


            Dim Delimiter As Char = "|"
            If StoreList.Contains("|") Then
                Delimiter = "|"
            Else
                Delimiter = ","
            End If

            Dim sb As New StringBuilder

            Dim StoreListAll As String() = StoreList.Split(Delimiter)

            For s As Integer = 0 To StoreListAll.Length - 1
                If s <> 0 Then
                    sb.Append(" union")
                End If
                sb.Append(" Select cast('" & StoreListAll(s) & "' as char) as StoreCode")
            Next
            Dim QueryStore As String = sb.ToString

            'tStart = Date.Now

            sb = New StringBuilder
            sb.Append(" Select x.StoreCode")
            sb.Append(" From (")
            sb.Append(QueryStore)
            sb.Append(" )x")
            sb.Append(" Left Join IndomaretStore i on (x.StoreCode = i.Code)")
            sb.Append(" Left Join MappingDepoStore depo on (x.StoreCode = depo.Store)")
            sb.Append(" Left Join MappingDcStore dc on (x.StoreCode = dc.Store)")
            sb.Append(" Where IfNull(i.Code,'') <> '' And ifnull(IfNull(depo.Depo,dc.DC),'') = @DCCode") 'hanya toko sesuai coverage
            Query = sb.ToString

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@DCCode", IIf(DCIGR <> DCCode, DCIGR, DCCode))

            Dim dtStore As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)
            If Not dtStore Is Nothing Then
                sb = New StringBuilder
                StoreList = ""
                For s As Integer = 0 To dtStore.Rows.Count - 1
                    If s <> 0 Then
                        sb.Append(Delimiter)
                    End If
                    sb.Append(dtStore.Rows(s).Item("StoreCode").ToString)
                Next
                StoreList = sb.ToString
            End If

            'tEnd = Date.Now
            'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
            'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)


            '=== Temp RAO
            Dim TempTableRao As String = "TempRaoDraft"

            Query = "Drop Temporary Table If Exists " & TempTableRao
            MCom.CommandText = Query
            MCom.Parameters.Clear()
            MCom.ExecuteNonQuery()

            'tStart = Date.Now

            Query = "Create Temporary Table " & TempTableRao & " ("
            Query &= " TrackNum varchar(20) not null default ''"
            Query &= " , Primary Key (`TrackNum`)"
            Query &= " )"
            Try
                MCom.CommandText = Query
                MCom.Parameters.Clear()
                MCom.ExecuteNonQuery()

                'tEnd = Date.Now
                'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                Result(1) &= "Error : " & ex.Message

                GoTo Skip
            End Try


            'tStart = Date.Now

            Query = "Insert Ignore Into " & TempTableRao & " (TrackNum)"
            Query &= " Select distinct TrackNum From RaoDraft Where Hub = @Hub"
            Try
                MCom.CommandText = Query

                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@Hub", Hub)

                MCom.ExecuteNonQuery()

                'tEnd = Date.Now
                'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                Result(1) &= "Error : " & ex.Message

                GoTo Skip
            End Try
            '=== Temp RAO


            MTrn = MCon.BeginTransaction()
            MCom.Transaction = MTrn

            Dim MaxTry As Short = 3
            Dim nTry As Short = 0

            'DCRequest
            'tStart = Date.Now

            sb = New StringBuilder
            sb.Append(" Insert Into DCRequest (")
            sb.Append(" DC, Hub, PickupNo, PickupTime, Store, `Status`, `Type`, UpdTime, UpdUser")
            If CustomRefNo <> "" Then
                sb.Append(" , `RefNo`")
            End If
            sb.Append(" ) values (")
            sb.Append(" @DCCode, @Hub, @PickupNo, @PickupTime, @StoreList, '0', @Tipe, now(), @User")
            If CustomRefNo <> "" Then
                sb.Append(" , @CustomRefNo")
            End If
            sb.Append(" ) On Duplicate Key Update UpdTime = now(), UpdUser = @User")
            sb.Append(" , PickupTime = @PickupTime, Store = @StoreList")
            Query = sb.ToString

            nTry = 0
UlangInsertDCRequest:
            Try
                nTry = nTry + 1

                MCom.CommandText = Query

                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@DCCode", DCCode)
                MCom.Parameters.AddWithValue("@Hub", Hub)
                MCom.Parameters.AddWithValue("@PickupNo", PickupNo)
                MCom.Parameters.AddWithValue("@PickupTime", Format(PickupTime, "yyyy-MM-dd HH:mm:ss"))
                MCom.Parameters.AddWithValue("@StoreList", StoreList.Replace("|", ","))
                MCom.Parameters.AddWithValue("@Tipe", Tipe)
                If CustomRefNo <> "" Then
                    MCom.Parameters.AddWithValue("@CustomRefNo", CustomRefNo)
                End If
                MCom.Parameters.AddWithValue("@User", User)

                MCom.ExecuteNonQuery()

                'tEnd = Date.Now
                'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

            Catch ex As Exception
                Dim e As New ClsError
                If nTry < MaxTry And ex.Message.ToUpper.Contains("DEADLOCK") Then
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, "Try " & nTry & vbCrLf & Query)

                    GoTo UlangInsertDCRequest
                Else
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                    Result(1) &= "Error : " & ex.Message

                    GoTo Skip
                End If
            End Try


            '=== Temp RackIn
            Dim TempTableRackIn As String = "TempRaiDraft"

            Query = "Drop Temporary Table If Exists " & TempTableRackIn
            MCom.CommandText = Query
            MCom.Parameters.Clear()
            MCom.ExecuteNonQuery()

            'tStart = Date.Now

            sb = New StringBuilder
            sb.Append("Create Temporary Table " & TempTableRackIn & "")
            sb.Append(" Select i.TrackNum, 'DCR' as `Status`, i.Rack, i.Dst, d.PickupNo")
            sb.Append(" , 'IDM' as Company1, 'DCI' as Opr1, d.PickupNo as OprID1")
            sb.Append(" , now() as TrackTime, 'DCI' as TrackUser, @Hub as TrackUserID")
            sb.Append(" From DCRequest d, RackIn i, DCTempRequest t")
            sb.Append(" Where d.Hub = @Hub And d.PickupNo = @PickupNo And d.Hub = i.Hub")
            If TrackNumList <> "" Then
                sb.Append(" And TrackNum in (" & ConvertNumListToSQL(TrackNumList.Replace(",", "|")) & ")")
            End If
            sb.Append(" And d.`Status` = '0'") 'belum WHO
            sb.Append(" And instr(d.Store,i.dst)") 'daftar toko yang di-request
            sb.Append(" And d.`Type` = t.`Type`")
            sb.Append(" And t.Hub = i.Hub And t.TrackNum = i.TrackNum") 'link ke table bantuan
            Query = sb.ToString
            Try
                MCom.CommandText = Query

                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@Hub", Hub)
                MCom.Parameters.AddWithValue("@PickupNo", PickupNo)

                MCom.ExecuteNonQuery()

                'tEnd = Date.Now
                'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                Result(1) &= "Error : " & ex.Message

                GoTo Skip
            End Try

            Query = "Alter Table " & TempTableRackIn
            Query &= " Add Primary Key (`TrackNum`)"
            Try
                MCom.CommandText = Query
                MCom.Parameters.Clear()
                MCom.ExecuteNonQuery()
            Catch ex As Exception
            End Try


            'tStart = Date.Now

            Query = "Delete r"
            Query &= " From `" & TempTableRackIn & "` r, `" & TempTableRao & "` o"
            Query &= " Where r.TrackNum = o.TrackNum"
            Try
                MCom.CommandText = Query
                MCom.Parameters.Clear()
                MCom.ExecuteNonQuery()

                'tEnd = Date.Now
                'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                Result(1) &= "Error : " & ex.Message

                GoTo Skip
            End Try
            '=== Temp RackIn


            'Insert Tracking
            'tStart = Date.Now

            sb = New StringBuilder
            sb.Append(" Insert Into Tracking (")
            sb.Append(" TrackNum, `Status`, AddInfo")
            sb.Append(" , Company1, Opr1, OprID1")
            sb.Append(" , TrackTime, TrackUser, TrackUserID")
            sb.Append(" ) Select TrackNum, 'DCR', @TrackingAddInfo")
            sb.Append(" , 'IDM', 'DCI', PickupNo")
            sb.Append(" , now(), 'DCI', @Hub")
            sb.Append(" From " & TempTableRackIn & "")
            Query = sb.ToString

            nTry = 0
UlangTracking:
            Try
                nTry = nTry + 1

                MCom.CommandText = Query

                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@Hub", Hub)
                MCom.Parameters.AddWithValue("@TrackingAddInfo", TrackingAddInfo)

                MCom.ExecuteNonQuery()

                'tEnd = Date.Now
                'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

            Catch ex As Exception
                Dim e As New ClsError
                If nTry < MaxTry And ex.Message.ToUpper.Contains("DEADLOCK") Then
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, "Try " & nTry & vbCrLf & Query)

                    GoTo UlangTracking
                Else
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                    Result(1) &= "Error : " & ex.Message

                    GoTo Skip
                End If
            End Try


            'buat draft RAO
            'tStart = Date.Now

            sb = New StringBuilder
            sb.Append(" Insert Ignore Into RaoDraft (")
            sb.Append(" Hub, PickupNo, Rack, TrackNum, Dst, `Status`, UpdTime, UpdUser")
            sb.Append(" ) Select")
            sb.Append(" @Hub, PickupNo, Rack, TrackNum, Dst, '0', now(), @User")
            sb.Append(" From " & TempTableRackIn & "")
            Query = sb.ToString

            nTry = 0
UlangRAODraft:
            Try
                nTry = nTry + 1

                MCom.CommandText = Query

                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@Hub", Hub)
                MCom.Parameters.AddWithValue("@User", User)

                MCom.ExecuteNonQuery()

                'tEnd = Date.Now
                'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

            Catch ex As Exception
                Dim e As New ClsError
                If nTry < MaxTry And ex.Message.ToUpper.Contains("DEADLOCK") Then
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, "Try " & nTry & vbCrLf & Query)

                    GoTo UlangRAODraft
                Else
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                    Result(1) &= "Error : " & ex.Message

                    GoTo Skip
                End If
            End Try


            If Tipe = "2" Then
                'Auto Picking Kurir Wilayah
                'tambah table bantuan untuk Daftar Antar Paket ke Toko di app Hub Del

                'tStart = Date.Now

                'nama table disamakan dengan proses KurirWilayahAntarTokoList
                Dim TbName As String = "RaoDraft_KurWil_" & Format(Date.Now, "yyMMdd")
                Dim TbNameDetail As String = "Detail" & TbName

                sb = New StringBuilder
                sb.Append("Create Table If Not Exists " & TbName & " Like RaoDraft")
                Query = sb.ToString
                Try
                    MCom.CommandText = Query
                    MCom.Parameters.Clear()
                    MCom.ExecuteNonQuery()
                Catch ex As Exception
                End Try

                sb = New StringBuilder
                sb.Append("Create Table if not exists " & TbNameDetail & " (")
                sb.Append("  Cluster varchar(50) Not NULL DEFAULT '' COMMENT 'Cluster'")
                sb.Append(" , Hub varchar(10) Not NULL DEFAULT '' COMMENT 'Kode Hub'")
                sb.Append(" , HubName varchar(50) Not NULL DEFAULT '' COMMENT 'Nama Hub'")
                sb.Append(" , PickupNo varchar(50) Not NULL DEFAULT '' COMMENT 'Pickup Number'")
                sb.Append(" , CreateDate datetime Not NULL DEFAULT '2000-12-31 00:00:00' COMMENT 'Tanggal dibuat'")
                sb.Append(" , PickupNoShort varchar(50) Not NULL DEFAULT '' COMMENT 'Pickup Number Pendek'")
                sb.Append(" , Jumlah Integer Not NULL DEFAULT 0 COMMENT 'Jumlah tracknum per pickup no'")
                sb.Append(" , Dst varchar(10) Not NULL DEFAULT '' COMMENT 'Kode toko tujuan'")
                sb.Append(" , DstName varchar(50) Not NULL DEFAULT '' COMMENT 'Nama toko'")
                sb.Append(" , DstLatitude varchar(50) Not NULL DEFAULT '' COMMENT 'Latitude'")
                sb.Append(" , DstLongitude varchar(50) Not NULL DEFAULT '' COMMENT 'Longitude'")
                sb.Append(" , DstPhone varchar(50) Not NULL DEFAULT '' COMMENT 'Telp Toko'")
                sb.Append(" , Tracknum varchar(50) Not NULL DEFAULT '' COMMENT 'AWB'")
                sb.Append(" , Status  smallint(5) Not NULL DEFAULT 0 COMMENT 'Status push'")
                sb.Append(" , RetryPush smallint(5) Not NULL DEFAULT 0 COMMENT 'Jumlah Percobaan push'")
                sb.Append(" , UpdTime datetime Not NULL DEFAULT '2000-12-31 00:00:00' COMMENT 'Tanggal ubah'")
                sb.Append(" , UpdUser varchar(50) Not NULL DEFAULT ''")
                sb.Append(" , PRIMARY KEY(`Hub`,`Cluster`,`Tracknum`)")
                sb.Append(" , KEY `i_Tracknum` (`Hub`, `Tracknum`)")
                sb.Append(" , KEY `i_cluster` (`cluster`)")
                sb.Append(" , KEY `i_Hub` (`Hub`)")
                sb.Append(" ) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Draft detail antar'")
                Query = sb.ToString
                Try
                    MCom.CommandText = Query
                    MCom.Parameters.Clear()
                    MCom.ExecuteNonQuery()
                Catch ex As Exception
                End Try


                sb = New StringBuilder
                sb.Append(" Insert Ignore Into " & TbName & " (")
                sb.Append(" Hub, PickupNo, Rack, TrackNum, Dst, `Status`, UpdTime, UpdUser")
                sb.Append(" ) Select @Hub, PickupNo, Rack, TrackNum, Dst, '0', now(), @User")
                sb.Append(" From " & TempTableRackIn & "")
                sb.Append(" Where PickupNo = @PickupNo")

                Query = sb.ToString

                nTry = 0

UlangRAODraftKurWil:
                Try
                    nTry = nTry + 1

                    MCom.CommandText = Query

                    MCom.Parameters.Clear()
                    MCom.Parameters.AddWithValue("@Hub", Hub)
                    MCom.Parameters.AddWithValue("@PickupNo", PickupNo)
                    MCom.Parameters.AddWithValue("@User", User)

                    MCom.ExecuteNonQuery()

                    'tEnd = Date.Now
                    'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                    'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

                Catch ex As Exception
                    Dim e As New ClsError
                    If nTry < MaxTry And ex.Message.ToUpper.Contains("DEADLOCK") Then
                        e.ErrorLog(AppName, AppVersion, Method, User, ex, "Try " & nTry & vbCrLf & Query)

                        GoTo UlangRAODraftKurWil
                    Else
                        e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                        Result(1) &= "Error : " & ex.Message

                        GoTo Skip
                    End If
                End Try


                'daftar antar ke toko, yang akan di-push ke sistem hub delivery
                sb = New StringBuilder
                sb.Append("Insert Ignore Into " & TbNameDetail & " (")
                sb.Append(" Cluster, Hub, HubName")
                sb.Append(" , PickupNo, CreateDate, PickupNoShort, Jumlah")
                sb.Append(" , Dst, DstName, DstLatitude, DstLongitude, DstPhone, TrackNum")
                sb.Append(" ) Select x.Cluster, d.Hub, hub.Alias as HubName")
                sb.Append(" , d.PickupNo, d.PickupTime as CreateDate")
                sb.Append(" , cast(case when length(r.PickupNo) > " & GlobalSjNumShortLength & "  then concat(left(r.PickupNo,4),'..', right(r.PickupNo," & (GlobalSjNumShortLength - 4) & ")) else r.PickupNo end as char) as PickupNoShort")
                sb.Append(" , ifnull((Select count(TrackNum) From " & TbName & " Where PickupNo = d.PickupNo group by PickupNo),0) as Jumlah")
                sb.Append(" , r.Dst, i.Name as DstName, i.Latitude as DstLatitude, i.Longitude as DstLongitude, i.Phone as DstPhone")
                sb.Append(" , r.TrackNum as TrackNum")
                sb.Append(" From " & TbName & " r")
                sb.Append(" Inner Join IndomaretStore_For_Report i on (r.dst = i.code)")
                sb.Append(" Inner Join DcRequest d on (r.pickupno = d.pickupno)")
                sb.Append(" Inner Join MstHub hub on (hub.Hub = d.Hub)")
                sb.Append(" Inner Join(")
                sb.Append(" Select Cluster, concat('KW-', Cluster, '-', DATE_FORMAT(curdate(), '%y%m%d')) as RefNo")
                sb.Append(" From KurirWilayahMappingCluster")
                sb.Append(" )x On (d.RefNo = x.RefNo)")
                sb.Append(" Where r.PickupNo = d.PickupNo And d.type = '2'")
                sb.Append(" And r.PickupNo = @PickupNo")
                sb.Append(" Group by r.tracknum")
                sb.Append(" Order By r.Dst")
                Query = sb.ToString

                nTry = 0
UlangRAODraftKurWilDetail:
                Try
                    nTry = nTry + 1

                    MCom.CommandText = Query

                    MCom.Parameters.Clear()
                    MCom.Parameters.AddWithValue("@PickupNo", PickupNo)

                    MCom.ExecuteNonQuery()

                    'tEnd = Date.Now
                    'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                    'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

                Catch ex As Exception
                    Dim e As New ClsError
                    If nTry < MaxTry And ex.Message.ToUpper.Contains("DEADLOCK") Then
                        e.ErrorLog(AppName, AppVersion, Method, User, ex, "Try " & nTry & vbCrLf & Query)

                        GoTo UlangRAODraftKurWilDetail
                    Else
                        e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                        Result(1) &= "Error : " & ex.Message

                        GoTo Skip
                    End If
                End Try


                Dim ParamNotif(3) As Object
                ParamNotif(0) = User
                ParamNotif(1) = Password
                ParamNotif(2) = "ANTAR_TOKO"
                ParamNotif(3) = PickupNo
                KurirWilayahPushNotifListInsert(AppName, AppVersion, ParamNotif)

            End If 'dari If Tipe = "2"


            If StoreInfoList <> "" Then

                'tStart = Date.Now

                'pakai urutan zona toko
                Dim StoreInfoListSplit As String() = StoreInfoList.Split("|")

                MCom.Parameters.Clear()

                sb = New StringBuilder

                sb.Append(" Insert Ignore Into DCRequestZone (")
                sb.Append(" Hub, PickupNo, Store, `Zone`, SeqNo, UpdTime, UpdUser")
                sb.Append(" ) Values")

                For i As Integer = 0 To StoreInfoListSplit.Length - 1

                    If i <> 0 Then
                        sb.Append(",")
                    End If

                    Dim StoreInfoListSplitSplit As String() = StoreInfoListSplit(i).Split(",")
                    '0 = kode toko, 1 = zona, 2 = nomor urut

                    Dim Store As String = StoreInfoListSplitSplit(0)
                    Dim Zone As String = StoreInfoListSplitSplit(1)
                    Dim SeqNo As String = StoreInfoListSplitSplit(2)

                    sb.Append(" (@Hub, @PickupNo, @Store" & i & ", @Zone" & i & ", @SeqNo" & i & ", now(), @User)")

                    MCom.Parameters.AddWithValue("@Store" & i, Store)
                    MCom.Parameters.AddWithValue("@Zone" & i, Zone)
                    MCom.Parameters.AddWithValue("@SeqNo" & i, SeqNo)

                Next
                Query = sb.ToString

                nTry = 0
UlangDCRequestZone:
                Try
                    nTry = nTry + 1

                    MCom.CommandText = Query

                    MCom.Parameters.AddWithValue("@Hub", Hub)
                    MCom.Parameters.AddWithValue("@PickupNo", PickupNo)
                    MCom.Parameters.AddWithValue("@User", User)

                    MCom.ExecuteNonQuery()

                    'tEnd = Date.Now
                    'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                    'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

                Catch ex As Exception
                    Dim e As New ClsError
                    If nTry < MaxTry And ex.Message.ToUpper.Contains("DEADLOCK") Then
                        e.ErrorLog(AppName, AppVersion, Method, User, ex, "Try " & nTry & vbCrLf & Query)

                        GoTo UlangDCRequestZone
                    Else
                        e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                        Result(1) &= "Error : " & ex.Message

                        GoTo Skip
                    End If
                End Try


            Else

                'tStart = Date.Now

                'pakai urutan dari database rack
                Query = "Select count(Table_Name) From Information_Schema.`Tables`"
                Query &= " Where Table_Schema = @MyDbName And Table_Name = @MstRack_Hub"
                MCom.CommandText = Query

                Try
                    MCom.Parameters.Clear()
                    MCom.Parameters.AddWithValue("@MyDbName", MyDbName)
                    MCom.Parameters.AddWithValue("@MstRack_Hub", "MstRack_" & Hub)

                    If MCom.ExecuteScalar > 0 Then 'cek table rack hub

                        Query = "Select Store, Rack as Zone, Name as SeqNo From MstRack_" & Hub
                        Query &= " Where Store in (" & ConvertNumListToSQL(StoreList.Replace(",", "|")) & ")"
                        Query &= " Order By Rack, length(Name), Name"
                        MCom.CommandText = Query
                        MCom.Parameters.Clear()

                        Dim MDa As New MySqlDataAdapter(MCom)

                        Dim dtHubRackZone As New DataTable
                        MDa.Fill(dtHubRackZone)

                        If dtHubRackZone Is Nothing = False Then
                            If dtHubRackZone.Rows.Count > 0 Then

                                StoreInfoList = ""
                                For i As Integer = 0 To dtHubRackZone.Rows.Count - 1
                                    If StoreInfoList <> "" Then
                                        StoreInfoList &= "|"
                                    End If
                                    StoreInfoList &= dtHubRackZone.Rows(i).Item("Store").ToString
                                    StoreInfoList &= "," & dtHubRackZone.Rows(i).Item("Zone").ToString
                                    StoreInfoList &= "," & dtHubRackZone.Rows(i).Item("SeqNo").ToString
                                Next

                                Dim StoreInfoListSplit As String() = StoreInfoList.Split("|")

                                MCom.Parameters.Clear()

                                sb = New StringBuilder

                                sb.Append(" Insert Ignore Into DCRequestZone (")
                                sb.Append(" Hub, PickupNo, Store, `Zone`, SeqNo, UpdTime, UpdUser")
                                sb.Append(" ) Values")

                                For i As Integer = 0 To StoreInfoListSplit.Length - 1

                                    If i <> 0 Then
                                        sb.Append(",")
                                    End If

                                    Dim StoreInfoListSplitSplit As String() = StoreInfoListSplit(i).Split(",")
                                    '0 = kode toko, 1 = zona, 2 = nomor urut

                                    Dim Store As String = StoreInfoListSplitSplit(0)
                                    Dim Zone As String = StoreInfoListSplitSplit(1)
                                    Dim SeqNo As String = StoreInfoListSplitSplit(2)

                                    sb.Append(" (@Hub, @PickupNo, @Store" & i & ", @Zone" & i & ", @SeqNo" & i & ", now(), @User)")

                                    MCom.Parameters.AddWithValue("@Store" & i, Store)
                                    MCom.Parameters.AddWithValue("@Zone" & i, Zone)
                                    MCom.Parameters.AddWithValue("@SeqNo" & i, SeqNo)

                                Next
                                Query = sb.ToString

                                nTry = 0
UlangHubRackZone:
                                Try
                                    nTry = nTry + 1

                                    MCom.CommandText = Query

                                    MCom.Parameters.AddWithValue("@Hub", Hub)
                                    MCom.Parameters.AddWithValue("@PickupNo", PickupNo)
                                    MCom.Parameters.AddWithValue("@User", User)

                                    MCom.ExecuteNonQuery()

                                    'tEnd = Date.Now
                                    'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                                    'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

                                Catch ex As Exception
                                    Dim e As New ClsError
                                    If nTry < MaxTry And ex.Message.ToUpper.Contains("DEADLOCK") Then
                                        e.ErrorLog(AppName, AppVersion, Method, User, ex, "Try " & nTry & vbCrLf & Query)

                                        GoTo UlangHubRackZone
                                    Else
                                        e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                                        Result(1) &= "Error : " & ex.Message

                                        GoTo Skip
                                    End If
                                End Try

                            End If 'dari If dtHubRackZone.Rows.Count > 0
                        End If

                    End If 'dari If cek table rack hub

                Catch ex As Exception
                    Dim e As New ClsError
                    e.DebugLog(MCon, AppName, AppVersion, Method, "Gagal DcRequestZone berdasarkan HubRack " & ex.Message & " " & ex.StackTrace, LogKeyword)

                End Try

            End If 'dari If StoreInfoList <> ""


            'delete dari table bantuan
            'tStart = Date.Now

            Query = "Delete i"
            Query &= " From DCTempRequest i, RaoDraft o"
            Query &= " Where o.Hub = @Hub And o.PickupNo = @PickupNo"
            Query &= " And o.Hub = i.Hub And o.TrackNum = i.TrackNum"
            Query &= " And i.`Type` = @Tipe"

            nTry = 0
UlangDeleteDCTempRequest:
            Try
                nTry = nTry + 1

                MCom.CommandText = Query

                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@Hub", Hub)
                MCom.Parameters.AddWithValue("@PickupNo", PickupNo)
                MCom.Parameters.AddWithValue("@Tipe", Tipe)

                MCom.ExecuteNonQuery()

                'tEnd = Date.Now
                'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

            Catch ex As Exception
                Dim e As New ClsError
                If nTry < MaxTry And ex.Message.ToUpper.Contains("DEADLOCK") Then
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, "Try " & nTry & vbCrLf & Query)

                    GoTo UlangDeleteDCTempRequest
                Else
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                    Result(1) &= "Error : " & ex.Message

                    GoTo Skip
                End If
            End Try

            Process = True

Skip:
            If Process Then
                Try
                    MTrn.Commit()
                Catch ex As Exception
                End Try

                'tampilkan dari draft RAO
                Dim TempTableResponse As String = "TempPickingFinal"

                Query = "Drop Temporary Table If Exists " & TempTableResponse
                MCom.CommandText = Query
                MCom.Parameters.Clear()
                MCom.ExecuteNonQuery()

                Query = "Create Temporary Table " & TempTableResponse & " ("
                Query &= " TrackNum varchar(20) not null default ''"
                Query &= " , Dst varchar(10) not null default ''"
                Query &= " , Primary Key (`TrackNum`)"
                Query &= " , Key i_Dst (`Dst`)"
                Query &= " )"
                Try
                    MCom.CommandText = Query
                    MCom.Parameters.Clear()
                    MCom.ExecuteNonQuery()
                Catch ex As Exception
                End Try

                sb = New StringBuilder
                sb.Append("Insert Ignore Into " & TempTableResponse & "")
                sb.Append(" (TrackNum, Dst)")
                sb.Append(" Select TrackNum, Dst")
                sb.Append(" From DCRequest d, RaoDraft o")
                sb.Append(" Where d.Hub = o.Hub And d.PickupNo = o.PickupNo")
                sb.Append(" And d.DC = @DCCode And d.PickupNo = @PickupNo")
                sb.Append(" And d.`Type` = @Tipe")
                Query = sb.ToString
                Try
                    MCom.CommandText = Query

                    MCom.Parameters.Clear()
                    MCom.Parameters.AddWithValue("@DCCode", DCCode)
                    MCom.Parameters.AddWithValue("@PickupNo", PickupNo)
                    MCom.Parameters.AddWithValue("@Tipe", Tipe)

                    MCom.ExecuteNonQuery()
                Catch ex As Exception
                End Try


                'tStart = Date.Now

                sb = New StringBuilder
                sb.Append(" Select x.Dst as Store, t.TrackNum, t.Length, t.Width, t.Height, t.Weight")
                sb.Append(" From `Transaction` t, `" & TempTableResponse & "` x")
                sb.Append(" Where t.TrackNum = x.TrackNum")
                sb.Append(" Order by x.Dst")
                Query = sb.ToString

                Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)

                'tEnd = Date.Now
                'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, Query, LogKeyword)

                If Not dtQuery Is Nothing Then

                    'tStart = Date.Now

                    Result(2) = ConvertDatatableToString(dtQuery)

                    'tEnd = Date.Now
                    'tInterval = DateDiff(DateInterval.Second, tStart, tEnd)
                    'x.DebugLog(AppName, AppVersion, "DC Picking Final Debug", "budil " & tInterval, "Convert Result(2)", LogKeyword)

                    Result(0) = "0"
                    Result(1) = ResponseOK

                End If

            Else
                Try
                    MTrn.Rollback()
                Catch ex As Exception
                End Try
            End If

            CreateResponseLog(AppName, AppVersion, Method, User, Result, PickupNo, LogKeyword)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Public Function DCPickingDraft(ByVal AppName As String, ByVal AppVersion As String _
        , ByVal DCCode As String, ByVal Param As Object()) As Object()

        DCCode = DCCode.ToUpper

        Dim Method As String = "DCPickingDraft"

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            Dim StoreList As String = Param(2).ToString.Trim.ToUpper

            Dim TrackNumList As String = Param(4).ToString.Trim.ToUpper

            Dim Tipe As String = "0" 'default = DRY

            Dim tTipe As String = ""
            Try
                tTipe = ("" & Param(3)).ToString.ToUpper.Trim
            Catch ex As Exception
                tTipe = ""
            End Try

            If tTipe <> "" Then
                If tTipe.Contains("ROTI") Then
                    Tipe = "1"
                End If
                If tTipe.Contains("KURWIL") Then
                    Tipe = "2"
                End If
            End If

            Dim MTrn As MySqlTransaction

            MCon = MasterMCon.Clone

            Dim MCom As New MySqlCommand("", MCon)
            MCon.Open()

            CreateRequestLog(AppName, AppVersion, Method, User, DCCode, Param)

            Dim Process As Boolean = False

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            '    Result(1) = UserOK(1)
            '    GoTo Skip
            'End If

            Dim CheckDCCode As String = GetDC(MCon, DCCode)
            If CheckDCCode = "" Then
                CheckDCCode = GetDepo(MCon, DCCode)
            End If

            If CheckDCCode = "" Then
                Result(1) = GetDCDepoNotFound(DCCode)
                GoTo Skip
            End If

            Dim Hub As String = GetMappingHubDC(MCon, DCCode, True)
            If Hub = "" Then
                Result(1) = GetMappingHubDCNotFound(DCCode)
                GoTo Skip
            End If

            If StoreList = "" Then
                Result(1) = "Input Daftar Toko"
                GoTo Skip
            End If

            Dim DCIGR As String = ""
            MCom.CommandText = "Select DC From MappingHubDC Where IGR = '" & DCCode & "'"
            Try
                DCIGR = ("" & MCom.ExecuteScalar).ToString.Trim.ToUpper
            Catch ex As Exception
                DCIGR = ""
            End Try

            If DCIGR = "" Then
                DCIGR = DCCode
            End If

            Dim Delimiter As Char = "|"
            If StoreList.Contains("|") Then
                Delimiter = "|"
            Else
                Delimiter = ","
            End If

            Dim sb As New StringBuilder

            Dim StoreListAll As String() = StoreList.Split(Delimiter)

            For s As Integer = 0 To StoreListAll.Length - 1
                If s <> 0 Then
                    sb.Append(" union")
                End If
                sb.Append(" Select cast('" & StoreListAll(s) & "' as char) as StoreCode")
            Next
            Dim QueryStore As String = sb.ToString

            sb = New StringBuilder
            sb.Append(" Select x.StoreCode")
            sb.Append(" From (")
            sb.Append(QueryStore)
            sb.Append(" )x")
            sb.Append(" Left Join IndomaretStore i on (x.StoreCode = i.Code)")
            sb.Append(" Left Join MappingDepoStore depo on (x.StoreCode = depo.Store)")
            sb.Append(" Left Join MappingDcStore dc on (x.StoreCode = dc.Store)")
            sb.Append(" Where IfNull(i.Code,'') <> '' And ifnull(IfNull(depo.Depo,dc.DC),'') = '" & IIf(DCIGR <> DCCode, DCIGR, DCCode) & "'") 'hanya toko sesuai coverage
            Query = sb.ToString

            Dim ObjSQL As New ClsSQL

            Dim dtStore As DataTable = ObjSQL.ExecDatatable(MCon, Query)
            If Not dtStore Is Nothing Then
                sb = New StringBuilder
                StoreList = ""
                For s As Integer = 0 To dtStore.Rows.Count - 1
                    If s <> 0 Then
                        sb.Append(Delimiter)
                    End If
                    sb.Append(dtStore.Rows(s).Item("StoreCode").ToString)
                Next
                StoreList = sb.ToString
            End If


            '=== Temp RAO
            Dim TempTableRao As String = "TempRaoDraft"

            Query = "Drop Temporary Table If Exists " & TempTableRao
            MCom.CommandText = Query
            MCom.ExecuteNonQuery()

            Query = "Create Temporary Table " & TempTableRao & " ("
            Query &= " TrackNum varchar(20) not null default ''"
            Query &= " , Primary Key (`TrackNum`)"
            Query &= " )"
            Try
                MCom.CommandText = Query
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                Result(1) &= "Error : " & ex.Message

                GoTo Skip
            End Try

            Query = "Insert Ignore Into " & TempTableRao
            Query &= " (TrackNum)"
            Query &= " Select distinct TrackNum From RaoDraft"
            Query &= " Where Hub = '" & Hub & "'"
            Try
                MCom.CommandText = Query
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                Result(1) &= "Error : " & ex.Message

                GoTo Skip
            End Try
            '=== Temp RAO


            '=== Temp RAI
            Dim TempTableRackIn As String = "TempRaiDraft"

            Query = "Drop Temporary Table If Exists " & TempTableRackIn
            MCom.CommandText = Query
            MCom.ExecuteNonQuery()

            sb = New StringBuilder
            sb.Append("Create Temporary Table " & TempTableRackIn & "")
            sb.Append(" Select")
            sb.Append(" '" & Hub & "' as Hub, i.TrackNum, i.Dst")
            sb.Append(" , '" & Tipe & "' as `Type`")
            sb.Append(" , now() as UpdTime, '" & User & "' as UpdUser")
            sb.Append(" From RackIn i")
            sb.Append(" Where i.Hub = '" & Hub & "'")
            If TrackNumList <> "" Then
                sb.Append(" And i.TrackNum in (" & ConvertNumListToSQL(TrackNumList.Replace(",", "|")) & ")")
            End If
            sb.Append(" And instr('" & StoreList.Replace("'", "") & "',i.Dst)") 'daftar toko yang di-request
            Query = sb.ToString
            Try
                MCom.CommandText = Query
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                Result(1) &= "Error : " & ex.Message

                GoTo Skip
            End Try

            Query = "Alter Table " & TempTableRackIn
            Query &= " Add Primary Key (`TrackNum`)"
            Try
                MCom.CommandText = Query
                MCom.ExecuteNonQuery()
            Catch ex As Exception
            End Try

            Query = "Delete r"
            Query &= " From `" & TempTableRackIn & "` r, `" & TempTableRao & "` o"
            Query &= " Where r.TrackNum = o.TrackNum"
            Try
                MCom.CommandText = Query
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                Result(1) &= "Error : " & ex.Message

                GoTo Skip
            End Try
            '=== Temp RAI


            MTrn = MCon.BeginTransaction()
            MCom.Transaction = MTrn

            Dim MaxTry As Short = 3
            Dim nTry As Short = 0

            'hapus dari table bantuan, antisipasi sudah pernah ada request untuk toko yang sama
            Query = "Delete From DCTempRequest "
            Query &= " Where Hub = '" & Hub & "'"
            Query &= " And instr('" & StoreList.Replace("'", "") & "',Dst)" 'daftar toko yang di-request
            Query &= " And `Type` = '" & Tipe & "'"

            nTry = 0
UlangDeleteDCTempRequest:
            Try
                nTry = nTry + 1

                MCom.CommandText = Query
                MCom.ExecuteNonQuery()

            Catch ex As Exception
                Dim e As New ClsError
                If nTry < MaxTry And ex.Message.ToUpper.Contains("DEADLOCK") Then
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, "Try " & nTry & vbCrLf & Query)

                    GoTo UlangDeleteDCTempRequest
                Else
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                    Result(1) &= "Error : " & ex.Message

                    GoTo Skip
                End If
            End Try



            'insert ke table bantuan
            sb = New StringBuilder
            sb.Append(" Insert Into DCTempRequest (")
            sb.Append(" Hub, TrackNum, Dst")
            sb.Append(" , `Type`")
            sb.Append(" , UpdTime, UpdUser")
            sb.Append(" ) Select")
            sb.Append(" '" & Hub & "', i.TrackNum, i.Dst")
            sb.Append(" , '" & Tipe & "'")
            sb.Append(" , now(), '" & User & "'")
            sb.Append(" From " & TempTableRackIn & " i")
            sb.Append(" On Duplicate Key Update Dst = values(Dst), UpdTime = now()")
            Query = sb.ToString

            nTry = 0
UlangInsertDCTempRequest:
            Try
                nTry = nTry + 1

                MCom.CommandText = Query
                MCom.ExecuteNonQuery()

            Catch ex As Exception
                Dim e As New ClsError
                If nTry < MaxTry And ex.Message.ToUpper.Contains("DEADLOCK") Then
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, "Try " & nTry & vbCrLf & Query)

                    GoTo UlangInsertDCTempRequest
                Else
                    e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

                    Result(1) &= "Error : " & ex.Message

                    GoTo Skip
                End If
            End Try


            Process = True

Skip:
            If Process Then
                Try
                    MTrn.Commit()
                Catch ex As Exception
                End Try

                'balikan ke DC
                sb = New StringBuilder
                sb.Append(" Select x.Dst as Store")
                sb.Append(" , t.TrackNum, t.Length, t.Width, t.Height, t.Weight")
                sb.Append(" From DCTempRequest x")
                sb.Append(" Inner Join `Transaction` t on (t.TrackNum = x.TrackNum)")
                sb.Append(" Where x.Hub = '" & Hub & "'")
                sb.Append(" And instr('" & StoreList.Replace("|", ",") & "',x.Dst)")
                sb.Append(" And x.`Type` = '" & Tipe & "'")
                sb.Append(" Order by x.Dst")
                Query = sb.ToString

                Dim dtQuery As DataTable = ObjSQL.ExecDatatable(MCon, Query)

                If Not dtQuery Is Nothing Then

                    Result(2) = ConvertDatatableToString(dtQuery)
                    Result(0) = "0"
                    Result(1) = ResponseOK

                End If

            Else
                Try
                    MTrn.Rollback()
                Catch ex As Exception
                End Try

            End If

            CreateResponseLog(AppName, AppVersion, Method, User, Result)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Private Function GetDC(ByVal MCon As MySqlConnection, ByVal DCCode As String) As String

        Dim DC As String = ""

        Try
            Dim ObjSQL As New ClsSQL
            'DC = "" & ObjSQL.ExecScalar(MCon, "Select `Code` From IndomaretDC Where `Code` = '" & DCCode & "'")

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@DCCode", DCCode)

            DC = "" & ObjSQL.ExecScalarWithParam(MCon, "Select `Code` From IndomaretDC Where `Code` = @DCCode", SqlParam)

        Catch ex As Exception
            DC = ""

        End Try

        Return DC

    End Function

    Private Function GetDepo(ByVal MCon As MySqlConnection, ByVal DepoCode As String) As String

        Dim Depo As String = ""

        Try
            Dim ObjSQL As New ClsSQL
            'Depo = "" & ObjSQL.ExecScalar(MCon, "Select `Code` From IndomaretDepo Where `Code` = '" & DepoCode & "'")

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@DepoCode", DepoCode)

            Depo = "" & ObjSQL.ExecScalarWithParam(MCon, "Select `Code` From IndomaretDepo Where `Code` = @DepoCode", SqlParam)

        Catch ex As Exception
            Depo = ""

        End Try

        Return Depo

    End Function

    Private Function GetDCDepoNotFound(ByVal DCCode As String) As String
        Return "DC atau Depo " & DCCode & " tidak ditemukan"
    End Function

    Private Function GetMappingDCStore(ByVal MCon As MySqlConnection, ByVal StoreCode As String) As String

        Dim DCCode As String = ""

        Try
            Dim ObjSQL As New ClsSQL
            'DCCode = "" & ObjSQL.ExecScalar(MCon, "Select DC From MappingDCStore Where Store = '" & StoreCode & "'")

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@StoreCode", StoreCode)

            DCCode = "" & ObjSQL.ExecScalarWithParam(MCon, "Select DC From MappingDCStore Where Store = @StoreCode", SqlParam)

        Catch ex As Exception
            DCCode = ""

        End Try

        Return DCCode

    End Function

    Public Function KurirWilayahPushNotifListInsert(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object

        Dim Method As String = "KurirWilayahPushNotifListInsert"

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection

        Dim e As New ClsError

        Dim LogKeyword As String = Format(Date.Now, "yyMMddHHmmss")
        Dim Query As String = ""

        Dim User As String = Param(0).ToString
        Try
            Dim Password As String = Param(1).ToString

            'Dim nLimit As Integer = 0
            'Try
            '    nLimit = CInt(Param(2))
            '    If IsNumeric(nLimit) = False Then
            '        nLimit = 0
            '    End If
            'Catch
            '    nLimit = 0
            'End Try
            'If nLimit < 1 Then
            '    nLimit = 20
            'End If

            Dim RedirectPage As String = ""
            Try
                RedirectPage = Param(2).ToString.Trim.ToUpper
            Catch ex As Exception
                RedirectPage = ""
            End Try

            Dim ProcessNo As String = ""
            Try
                ProcessNo = Param(3).ToString.Trim
            Catch ex As Exception
                ProcessNo = ""
            End Try

            MCon = MasterMCon.Clone
            MCon.Open()

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            If UserOK(0) <> "0" Then
                Result(1) = UserOK(1)
                GoTo Skip
            End If

            If RedirectPage = "" Then
                Result(1) = "Input RedirectPage"
                GoTo Skip
            End If

            CreateRequestLog(AppName, AppVersion, Method, User, "", Param, LogKeyword)

            Dim ObjSQL As New ClsSQL

            'proses hapus table secara berkala dengan proses KurirWilayahDeleteAntarTokoList
            Dim TbName As String = "Notif_KurWil_" & Format(Date.Now, "yyMMdd")

            Query = "Create Table if not exists " & TbName & " ("
            Query &= "  NotifId int Not NULL AUTO_INCREMENT COMMENT 'Notification ID'"
            Query &= " ,User varchar(50) Not NULL DEFAULT '' COMMENT 'NIK Kurir'"
            Query &= " ,Name varchar(50) Not NULL DEFAULT '' COMMENT 'Nama Kurir'"
            Query &= " ,nTracknum int Not NULL DEFAULT 0 COMMENT 'Jumlah AWB'"
            Query &= " ,Title varchar(255) Not NULL DEFAULT '' COMMENT 'Judul Notif'"
            Query &= " ,Description varchar(500) Not NULL DEFAULT '' COMMENT 'Desc Notif'"
            Query &= " ,RedirectPage varchar(20) Not NULL DEFAULT 0 COMMENT 'Page App'"
            Query &= " ,Status  smallint(5) Not NULL DEFAULT 0 COMMENT 'Status push'"
            Query &= " ,RetryPush smallint(5) Not NULL DEFAULT 0 COMMENT 'Jumlah Percobaan push'"
            Query &= " ,AddTime datetime Not NULL DEFAULT '2000-12-31 00:00:00' COMMENT 'Tanggal tambah'"
            Query &= " ,AddUser varchar(50) Not NULL DEFAULT ''"
            Query &= " ,UpdTime datetime DEFAULT NULL COMMENT 'Tanggal ubah'"
            Query &= " ,UpdUser varchar(50) Not NULL DEFAULT ''"
            Query &= " ,PRIMARY KEY(`NotifId`)"
            Query &= " ,KEY `i_Status` (`Status`)"
            Query &= " ,KEY `i_AddTime` (`AddTime`)"
            Query &= " ,KEY `i_RedirectPage` (`RedirectPage`)"
            Query &= " ,KEY `i_User` (`User`)"
            Query &= " ) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Draft Push Notif ke AppHubDel'"
            ObjSQL.ExecNonQueryWithParam(MCon, Query, Nothing)

            If RedirectPage = "HKPHBK" Then

                'perlu selesaikan HKP yang belum HBK

                Query = " insert ignore into " & TbName
                Query &= " (User, Name, nTracknum, Title, Description, RedirectPage, AddTime, AddUser)"
                Query &= " Select l.User, min(ucase(l.Name)) As Name, count(h.TrackNum) As nTrackNum"
                Query &= " , 'Antar Paket' as Title, left(concat('Halo ', ucase(l.name), ' (', l.`user`, ') ', ', ada ', count(h.TrackNum) ,' paket yang menunggu untuk diselesaikan (HBK)'), 500) as Description"
                Query &= " , @RedirectPage as RedirectPage, now(), @User"
                Query &= " From trackinghistory h"
                Query &= " inner Join mstlogin l on ("
                Query &= "  l.User = h.OprId2 And l.`Type` in ('HCR') "
                Query &= "  and l.`User` not like 'SUPDEL%' "
                Query &= "  and l.`User` not like 'SUPINSDEL%' "
                Query &= "  and l.`User` not like 'SUPOGW%' "
                Query &= " )"
                Query &= " where h.status in ('HDV','HOV') "
                'Query &= " and h.tracktime between concat(date_add(curdate(), interval -7 day),' 00:00:00') and concat(date_add(curdate(), interval -1 day),' 23:59:59') "
                Query &= " and h.tracktime between concat(date_add(curdate(), interval -7 day),' 00:00:00') and now() "
                Query &= " Group by l.User"
                'Query &= " limit " & nLimit

            ElseIf RedirectPage = "ANTAR_TOKO" Then

                'Kurwil perlu Antar ke Toko

                'nama table disamakan dengan proses DcPickingFinal
                Dim TbNameDetail As String = "DetailRaoDraft_KurWil_" & Format(Date.Now, "yyMMdd")

                Query = " Select Count(*) From information_schema.`tables`"
                Query &= " Where table_schema = '" & MyDbName & "' And table_name = '" & TbNameDetail & "'"
                If ObjSQL.ExecScalarWithParam(MCon, Query, Nothing) < 1 Then
                    Result(1) = "Tidak ada table " & TbNameDetail
                    GoTo Skip
                End If

                Query = " insert ignore into " & TbName
                Query &= " (User, Name, nTracknum, Title, Description, RedirectPage, AddTime, AddUser)"
                Query &= " Select l.`User`, min(ucase(l.Name)) As Name, count(r.TrackNum) As nTrackNum"
                Query &= " , 'Antar Paket ke Toko' as Title, left(concat('Halo ', ucase(l.name), ' (', l.`User`, ') ', ', ada ', count(r.TrackNum) ,' paket yang menunggu untuk diantar ke Toko (', r.Cluster, ')'), 500) as Description"
                Query &= " , @RedirectPage as RedirectPage, now(), @User"
                Query &= " From " & TbNameDetail & " r "
                Query &= " inner Join KurirWilayahMappingCluster d on (d.cluster = r.cluster)"
                Query &= " inner Join mstlogin l on (l.`User` = d.userid  And l.`Type` in ('HCR') )"
                Query &= " Where True"
                If ProcessNo <> "" Then
                    Query &= " And r.PickupNo = @ProcessNo"
                End If
                Query &= " Group by l.`User`, r.Cluster"
                'Query &= " limit  " & nLimit

            ElseIf RedirectPage = "JEMPUT_TOKO" Then

                'Kurwil perlu Jemput dari Toko

                'nama table disamakan dengan proses SuratTugasKurirIppProcess
                Dim TbNameDetail As String = "DetailJemputDraft_KurWil_" & Format(Date.Now, "yyMMdd")

                Query = " Select Count(*) From information_schema.`tables`"
                Query &= " Where table_schema = '" & MyDbName & "' And table_name = '" & TbNameDetail & "'"
                If ObjSQL.ExecScalarWithParam(MCon, Query, Nothing) < 1 Then
                    Result(1) = "Tidak ada table " & TbNameDetail
                    GoTo Skip
                End If

                Query = " insert ignore into " & TbName
                Query &= " (User, Name, nTracknum, Title, Description, RedirectPage, AddTime, AddUser)"
                Query &= " Select  l.`User`, min(ucase(l.Name)) As Name, count(r.TrackNum) As nTrackNum"
                Query &= " , 'Jemput Paket di Toko' as Title, left(concat('Halo ', ucase(l.name), ' (', l.`user`, ') ', ', ada ', count(r.TrackNum) ,' paket yang menunggu untuk dijemput di Toko (', r.Cluster, ')'), 500) as Description"
                Query &= " , @RedirectPage as RedirectPage, now(), @User"
                Query &= " From " & TbNameDetail & " r "
                Query &= " inner join KurirWilayahMappingCluster d on (d.cluster = r.cluster)"
                Query &= " inner Join mstlogin l on (l.`User` = d.userid  And l.`Type` in ('HCR') )"
                Query &= " Where True"
                If ProcessNo <> "" Then
                    Query &= " And r.SjNumH = @ProcessNo"
                End If
                Query &= " Group by l.`User`, r.Cluster"
                'Query &= " limit " & nLimit

            ElseIf RedirectPage = "JEMPUT_SAMEDAY" Then

                Query = " insert ignore into " & TbName
                Query &= " (User, Name, nTracknum, Title, Description, RedirectPage, AddTime, AddUser)"
                Query &= " Select l.User, min(ucase(l.Name)) As Name, count(t.TrackNum) As nTrackNum"
                Query &= " , 'Jemput Sameday' as Title, left(concat('Halo ', ucase(l.name), ' (', l.`user`, ') ', ', ada ', count(t.TrackNum) ,' paket yang menunggu untuk dijemput (', m.Cluster,')'), 500) as Description"
                Query &= " , @RedirectPage as RedirectPage, now(), @User"
                Query &= " from `Transaction` t"
                Query &= " inner join MstService v on (v.Service = t.ServiceType and ucase(v.Name) like 'SAMEDAY%DOOR')"
                Query &= " inner join TrackingHistory h on (h.TrackNum = t.TrackNum and h.`Status` = 'NEW')"
                Query &= " inner join KurirWilayahDistrictCoverage w on (w.District = t.ShDistrict)"
                Query &= " inner join KurirWilayahDistrictMappingCluster m on (m.Cluster = w.Cluster)"
                Query &= " inner join MstLogin l on (l.`user` = m.UserId and l.`type`= 'HCR')"
                Query &= " where t.TrackNum = t.oTrackNum "
                Query &= " and t.AddTime >= date_add(curdate(), interval -7 day)"
                Query &= " group by l.User, m.Cluster"
                'Query &= " limit " & nLimit

            Else
                Result(1) = "RedirectPage " & RedirectPage & " tidak dikenali"
                GoTo Skip

            End If

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@RedirectPage", RedirectPage)
            SqlParam.Add("@User", User)
            If ProcessNo <> "" Then
                SqlParam.Add("@ProcessNo", ProcessNo)
            End If

            ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

            Result(0) = "0"
            Result(1) = ResponseOK
            Result(2) = ""

Skip:

            CreateResponseLog(AppName, AppVersion, Method, User, Result, , LogKeyword)

        Catch ex As Exception
            e.ErrorLog(wsAppName, wsAppVersion, Method, User, ex, Query)

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

    Public Function GetExpeditionList(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim MCon2 As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            Dim ReqAddInfo As Boolean = False
            Try
                ReqAddInfo = (Param(4).ToString.Trim = "1")
            Catch ex As Exception
                ReqAddInfo = False
            End Try

            Dim UseJalurEkspedisi As Boolean = False

            MCon = MasterMCon.Clone
            MCon2 = Master2MCon.Clone

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'Else

            Dim ObjSQL As New ClsSQL

            Query = "Select `Account` as Value"
            Query &= " , cast(concat(case when IfNull(Alias,'') <> '' then Alias else `Name` end, ' (',right(Account,3) ,')') as char) As Display"
            If ReqAddInfo Then
                Query &= " , AddInfo"
            End If
            Query &= " From `account_test` Where 1=1"
            'Query &= " And curdate() between ActiveDate and IfNull(InactiveDate,curdate())"
            Query &= " And `Type` = '3'" 'penanda account type = Ekspedisi
            Query &= " Order By Alias"

            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon2, Query, Nothing)

            If Not dtQuery Is Nothing Then

                Dim dtTemp As New DataTable

                dtTemp = dtQuery.Clone

                If UseJalurEkspedisi = False Then
                    If ReqAddInfo Then
                        dtTemp.Rows.Add(New String() {"", "", ""})
                    Else
                        dtTemp.Rows.Add(New String() {"", ""})
                    End If
                Else
                    If ReqAddInfo Then
                        dtTemp.Rows.Add(New String() {"", "", "", "", ""})
                    Else
                        dtTemp.Rows.Add(New String() {"", "", "", ""})
                    End If
                End If

                For Each drow As DataRow In dtQuery.Rows
                    dtTemp.ImportRow(drow)
                Next

                Result(2) = ConvertDatatableToString(dtTemp)
                Result(1) = ResponseOK
                Result(0) = "0"

            End If

            'End If

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetExpeditionList", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

        Finally
            If MCon.State <> ConnectionState.Closed Then
                MCon.Close()
            End If
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            Try
                MCon2.Close()
            Catch ex As Exception
            End Try
            Try
                MCon2.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result

    End Function

    <WebMethod()>
    Public Function GetHubExpeditionList(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim MCon2 As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString
            Dim Hub As String = Param(2).ToString.Trim

            Dim ProcessType As String = ""
            Try
                ProcessType = Param(3).ToString.Trim.ToUpper
            Catch ex As Exception
                ProcessType = ""
            End Try

            Dim ReqAddInfo As Boolean = False
            Try
                ReqAddInfo = (Param(4).ToString.Trim = "1")
            Catch ex As Exception
                ReqAddInfo = False
            End Try

            Dim ReferenceNo As String = ""
            Try
                ReferenceNo = Param(5).ToString.Trim.ToUpper 'nomor referensi, misalnya nomor resi konsol (web hub)
            Catch ex As Exception
                ReferenceNo = ""
            End Try

            MCon = MasterMCon.Clone
            MCon2 = Master2MCon.Clone

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'GoTo Skip
            'End If

            If Hub = "" Or ProcessType = "" Then
                Result(1) = "Input Hub dan Jenis Proses"
                GoTo Skip
            End If

            ProcessType = ConvertNumListToSQL(ProcessType)

            Dim ObjSQL As New ClsSQL

            SqlParam = New Dictionary(Of String, String)

            If ProcessType.Contains("RESIKONSOL") And ReferenceNo <> "" Then
                Query = " Select a.`account_test` as Value"
                Query &= " , cast(concat(case when IfNull(a.Alias,'') <> '' then a.Alias else a.`Name` end, ' (',right(a.Account,3) ,')') as char) As Display"
                If ReqAddInfo Then
                    Query &= " , a.AddInfo"
                End If
                Query &= " from resikonsol_data d"
                Query &= " inner join resikonsol_datarute r on (r.nomor = d.nomor)"
                Query &= " inner join account a on (a.account = d.expedition)"
                Query &= " Where d.Nomor = @ReferenceNo and r.Hub = @Hub"

                SqlParam.Add("@Hub", Hub)
                SqlParam.Add("@ReferenceNo", ReferenceNo)
            Else
                Query = " Select a.`Account` as Value"
                Query &= " , cast(concat(case when IfNull(a.Alias,'') <> '' then a.Alias else a.`Name` end, ' (',right(a.Account,3) ,')') as char) As Display"
                If ReqAddInfo Then
                    Query &= " , a.AddInfo"
                End If
                Query &= " From MstHubExpedition m"
                Query &= " Inner Join `account_test` a on (a.Account = m.Expedition)"
                Query &= " Where m.Hub = @Hub"
                Query &= " and m.ProcessType in (" & ProcessType & ")"
                'Query &= " And curdate() between a.ActiveDate and IfNull(a.InactiveDate,curdate())"
                Query &= " Order By a.Alias"

                SqlParam.Add("@Hub", Hub)
            End If

            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon2, Query, SqlParam)

            If dtQuery Is Nothing Then
                Result(1) = "Gagal query"
                GoTo Skip
            End If

            Dim dtTemp As New DataTable

            dtTemp = dtQuery.Clone

            If ReqAddInfo Then
                dtTemp.Rows.Add(New String() {"", "", ""})
            Else
                dtTemp.Rows.Add(New String() {"", ""})
            End If

            For Each drow As DataRow In dtQuery.Rows
                dtTemp.ImportRow(drow)
            Next

            Result(2) = ConvertDatatableToString(dtTemp)
            Result(0) = "0"
            Result(1) = ResponseOK

Skip:

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "GetHubExpeditionList", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

        Finally
            If MCon.State <> ConnectionState.Closed Then
                MCon.Close()
            End If
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            Try
                MCon2.Close()
            Catch ex As Exception
            End Try
            Try
                MCon2.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result

    End Function

    <WebMethod()>
    Public Function SetOtherExpeditionAWB(ByVal AppName As String, ByVal AppVersion As String _
    , ByVal OriCode As String, ByVal Param As Object()) As Object()

        Dim Method As String = "SetOtherExpeditionAWB"

        OriCode = OriCode.ToUpper

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection

        Dim Query As String = ""

        Dim LogKeyword As String = Format(Date.Now, "yyMMddHHmmss")

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString
            Dim OtherExpeditionList As String = Param(2).ToString.ToUpper.Trim 'ConsNum,SjNum,OtherExp,OtherExpAWB,TrackNum,WeightKG|...

            Dim oExpedition As String = ""

            Dim AWB3PLBCKDTE As Integer = 0

            Try
                oExpedition = Param(3).ToString.Trim.ToUpper
            Catch ex As Exception
                oExpedition = ""
            End Try
            Dim oAWB As String = ""
            Try
                oAWB = Param(4).ToString.Trim.ToUpper
            Catch ex As Exception
                oAWB = ""
            End Try
            Dim oOri As String = ""
            Try
                oOri = Param(5).ToString.Trim.ToUpper
            Catch ex As Exception
                oOri = ""
            End Try
            Dim oDst As String = ""
            Try
                oDst = Param(6).ToString.Trim.ToUpper
            Catch ex As Exception
                oDst = ""
            End Try
            Dim oWeight As String = ""
            Try
                oWeight = Param(7).ToString.Trim.ToUpper
            Catch ex As Exception
                oWeight = ""
            End Try
            Dim oAWBDate As String = ""
            Try
                oAWBDate = Param(8).ToString.Trim.ToUpper
            Catch ex As Exception
                oAWBDate = ""
            End Try

            Dim oAWBDCost As Double = -1
            Try
                If Not Param(9) Is Nothing Then
                    If IsNumeric(Param(9)) Then
                        oAWBDCost = CDbl(Param(9))
                        If oAWBDCost = 0 Then
                            oAWBDCost = -1
                        End If
                    End If
                End If
            Catch ex As Exception
                oAWBDCost = -1
            End Try


            Dim AutoAdjustExpAwbDate As Boolean = False
            Try
                AutoAdjustExpAwbDate = (Param(10).ToString = "1")
            Catch ex As Exception
                AutoAdjustExpAwbDate = False
            End Try


            Dim IsTrackNum3plByIpp As Boolean = False
            Try
                IsTrackNum3plByIpp = (Param(11).ToString = "1")
            Catch ex As Exception
                IsTrackNum3plByIpp = False
            End Try

            Dim TrackNum3plByIpp As String = ""


            'menentukan biaya ekspedisi yang digunakan
            'CON = Serah Ekspedisi
            'AWB = Serah Antar Alamat
            'PUE = Jemput Ekspedisi
            'PUP = Jemput Partner
            'FLT = Sewa Armada
            Dim ExpenseType As String = ""
            Try
                ExpenseType = Param(12).ToString.Trim.ToUpper
            Catch ex As Exception
                ExpenseType = ""
            End Try
            If ExpenseType = "" Then
                ExpenseType = "CON"
            End If
            LogKeyword = ExpenseType & LogKeyword

            Dim ExpenseTbName As String = ExpeditionExpenseTableName(ExpenseType)

            Dim oDstCity As String = "" 'untuk ExpenseType AWB
            Try
                oDstCity = Param(13).ToString.Trim.ToUpper
            Catch ex As Exception
                oDstCity = ""
            End Try

            'untuk ExpenseType PUP dan FLT
            Dim oQtyVehicle As Integer = 0
            Try
                oQtyVehicle = CInt(Param(14))
            Catch ex As Exception
                oQtyVehicle = 0
            End Try

            'untuk ExpenseType PUP
            Dim oQtyPickup As Integer = 0
            Try
                oQtyPickup = CInt(Param(15))
            Catch ex As Exception
                oQtyPickup = 0
            End Try

            If ExpenseType = "PUP" Then
                oWeight = oQtyVehicle * oQtyPickup
            End If

            'untuk ExpenseType FLT
            Dim oVehicleType As String = ""
            Try
                oVehicleType = Param(16).ToString.Trim.ToUpper
            Catch ex As Exception
                oVehicleType = ""
            End Try

            Dim UserId As String = ""
            Try
                UserId = Param(17).ToString.Trim.ToUpper
            Catch ex As Exception
                UserId = ""
            End Try
            If UserId = "" Then
                UserId = User & "_NA"
            End If

            Dim MTrn As MySqlTransaction

            MCon = MasterMCon.Clone

            Dim MCom As New MySqlCommand("", MCon)
            MCon.Open()

            CreateRequestLog(AppName, AppVersion, Method, User, OriCode, Param, LogKeyword)

            MTrn = MCon.BeginTransaction()
            MCom.Transaction = MTrn

            Dim Process As Boolean = False

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'GoTo Skip
            'End If

            If GetHub(MCon, OriCode) = "" Then
                If GetStation(MCon, OriCode) = "" Then
                    Result(1) = GetStationOrHubNotFound(OriCode)
                    GoTo Skip
                End If
            End If

            If OtherExpeditionList = "" Then
                Result(1) = "Input Cons, Ekspedisi dan nomor Tracking"
                'GoTo Skip
            End If

            If oAWBDate = "" Or (IsDate(oAWBDate) = False) Then
                Result(1) = "Input Tanggal AWB dengan benar"
                'GoTo Skip
            End If

            Query = "select cast(curdate() as char)"
            MCom.CommandText = Query
            Dim SqlCurDate As String = MCom.ExecuteScalar

            'jagaan tanggal awb 3pl
            Query = "Select case when date_add('" & Format(CDate(SqlCurDate), "yyyy-MM-dd") & "', interval -" & AWB3PLBCKDTE & " day) <= '" & Format(CDate(oAWBDate), "yyyy-MM-dd") & "' then '1' else '0' end"
            MCom.CommandText = Query
            If MCom.ExecuteScalar < 1 Then
                If AutoAdjustExpAwbDate Then
                    oAWBDate = Format(DateAdd(DateInterval.Day, -1 * AWB3PLBCKDTE, CDate(SqlCurDate)), "yyyy-MM-dd")
                Else
                    Result(1) = "Tanggal AWB 3PL max. " & AWB3PLBCKDTE & " hari yang lalu (" & Format(DateAdd(DateInterval.Day, -1 * AWB3PLBCKDTE, CDate(SqlCurDate)), "yyyy-MM-dd") & ")"
                    'GoTo Skip
                End If
            End If

            'jagaan jumlahpickup dan kendaraan
            Query = "Select `Value` From const Where Key1 = 'MaxLimitNumber'"
            MCom.CommandText = Query
            Dim MaxNumber As String = MCom.ExecuteScalar

            If oQtyVehicle > CInt(MaxNumber) Then
                Result(1) = "Jumlah digit Kendaraan hanya bisa 1 digit angka!"
                GoTo Skip
            ElseIf oQtyPickup > CInt(MaxNumber) Then
                Result(1) = "Jumlah digit PickUp hanya bisa 1 digit angka!"
                GoTo Skip
            Else
            End If

            If ExpenseType = "AWB" Then 'kiriman to customer

                If oDstCity <> "" Then
                    oDst = oDstCity
                Else
                    Dim eDstCity As Integer = 0

                    Dim eSjNum As String = OtherExpeditionList.Split(",")(1)
                    Dim eTrackNum As String = OtherExpeditionList.Split(",")(4)
                    Query = "Select DstCity From SuratJalanD Where SjNum = '" & eSjNum & "' And TrackNum = '" & eTrackNum & "'"
                    MCom.CommandText = Query
                    Try
                        eDstCity = MCom.ExecuteScalar
                        If IsNumeric(eDstCity) = False Then
                            eDstCity = -1
                        End If
                    Catch ex As Exception
                        eDstCity = -1
                    End Try
                    If eDstCity < 0 Then
                        eDstCity = 0
                    End If

                    If eDstCity = 0 Then
                        Query = "Select CoCityCode From Transaction Where TrackNum = '" & eTrackNum & "'"
                        MCom.CommandText = Query
                        Try
                            eDstCity = MCom.ExecuteScalar
                            If IsNumeric(eDstCity) = False Then
                                eDstCity = -1
                            End If
                        Catch ex As Exception
                            eDstCity = -1
                        End Try
                        If eDstCity < 0 Then
                            eDstCity = 0
                        End If
                    End If

                    If eDstCity = 0 Then
                        Result(1) = "Gagal menentukan Kode Kota Tujuan"
                        GoTo Skip
                    End If

                    oDst = eDstCity
                End If

            End If 'dari If ExpenseType = "AWB" Then


            If ExpenseType = "PUP" Then
                If oQtyVehicle < 1 Or oQtyPickup < 1 Then
                    Result(1) = "Input Jumlah Kendaraan dan Jumlah Pickup"
                    GoTo Skip
                End If

                'supaya tidak gagal validasi di fungsi OtherExpeditionExpenseUpdate
                oOri = OriCode
                oDst = OriCode
            End If 'dari If ExpenseType = "PUP"


            If ExpenseType = "FLT" Then
                If oVehicleType = "" Then
                    Result(1) = "Input Jenis Kendaraan"
                    GoTo Skip
                End If
                If oQtyVehicle < 1 Then
                    Result(1) = "Input Jumlah Kendaraan"
                    GoTo Skip
                End If
            End If


            'cek harga 3PL sudah di-setting
            If OtherExpeditionDefault.Contains(oExpedition & ".") = False Then
                Query = "Select count(Account) From " & ExpenseTbName
                Query &= " Where Account = '" & oExpedition & "'"
                Select Case ExpenseType
                    Case "AWB"
                        Query &= " And Ori = '" & oOri & "' And Dst = '" & oDst & "'" '-- oDst = Kode Kota Tujuan
                    Case "PUE"
                        Query &= " And Dst = '" & oDst & "'" '-- oDst = Kode Hub yang proses
                    Case "PUP"
                        Query &= " And Ori = '" & oOri & "'" '-- oOri = Kode Hub yang proses
                    Case "FLT"
                        Query &= " And Ori = '" & oOri & "' And Dst = '" & oDst & "'" '-- Ori dan Dst = Kode Kota
                    Case Else
                        'CON
                        Query &= " And Ori = '" & oOri & "' And Dst = '" & oDst & "'" '-- oDst = Kode Hub Tujuan
                End Select
                Query &= " and curdate() between ActiveDate and ifnull(InActiveDate,curdate())"
                MCom.CommandText = Query
                If MCom.ExecuteScalar < 1 Then
                    If oAWBDCost > 0 Then
                        'tidak perlu cek Master Biaya Ekspedisi
                    Else
                        Result(1) = "Belum ada harga 3PL (" & oExpedition & "." & oOri & "." & oDst & ")"
                        GoTo Skip
                    End If
                End If
            End If

            'awb 3pl di-generate oleh IPP
            If oAWB = "" And IsTrackNum3plByIpp Then

                'bila ada perubahan, cek juga fungsi ExpeditionPickupProcess (proses Jemput Ekspedisi)
                'format nomor resi 3pl versi ipp
                '3 digit account ekspedisi
                '2 digit tahun
                '1 digit huruf pengganti bulan
                '4 digit running number

                Dim DateNow As Date = Date.Now

                Dim TrackNum3plFormat As String = Right(oExpedition, 3) & Format(DateNow, "yy") & ConvertMonthToAlphabet(Month(DateNow))

                Query = "Insert Into OtherExpeditionIppTracknum ("
                Query &= " Account, sYear, sMonth, SeqNo, TrackNum"
                Query &= " , Hub, Keyword, UpdTime, UpdUser"
                Query &= " ) Select a.Account, a.sYear, a.sMonth"
                Query &= " , ifnull(o.SeqNo,0) + 1 as SeqNo"
                Query &= " , ucase( concat( '" & TrackNum3plFormat & "', lpad( ifnull(o.SeqNo,0) + 1 ,4,'0') ) )"
                Query &= " , '" & OriCode & "', '" & LogKeyword & "', now(), '" & UserId & "'"
                Query &= " From ("
                Query &= "   Select Account, '" & Year(DateNow) & "' as sYear, '" & Month(DateNow) & "' as sMonth"
                Query &= "   From Account Where Account = '" & oExpedition & "'"
                Query &= " )a"
                Query &= " Left Join OtherExpeditionIppTracknum o on (o.Account = a.Account and o.sYear = a.sYear and o.sMonth = a.sMonth)"
                Query &= " Order By o.SeqNo Desc Limit 1"
                Try
                    MCom.CommandText = Query
                    MCom.ExecuteNonQuery()
                Catch ex As Exception
                    Dim e As New ClsError
                    e.ErrorLog(MCon, AppName, AppVersion, Method, OriCode, ex, Query, LogKeyword)

                    Result(1) &= "Error : " & ex.Message
                    GoTo Skip
                End Try

                Query = "Select TrackNum From OtherExpeditionIppTracknum"
                Query &= " Where Account = '" & oExpedition & "' And sYear = '" & Year(DateNow) & "' And sMonth = '" & Month(DateNow) & "'"
                Query &= " And Hub = '" & OriCode & "' And Keyword = '" & LogKeyword & "'"
                MCom.CommandText = Query
                TrackNum3plByIpp = ("" & MCom.ExecuteScalar()).ToString.Trim.ToUpper
                If TrackNum3plByIpp = "" Then
                    Result(1) = "Gagal generate Nomor Resi 3PL versi IPP"
                    GoTo Skip
                End If

                oAWB = TrackNum3plByIpp

            End If 'dari If oAWB = "" And Awb3plByIpp


            'validasi berat paket 3pl
            If oWeight <> "" And oExpedition <> "" And ExpenseType <> "PUP" And ExpenseType <> "FLT" Then

                Query = "Select ifnull(`Value`,0) From Const"
                Query &= " Where Key1 = 'MaxWeight3PL' And Key2 in (@oExpedition, '', '*')" 'dalam Gram
                Query &= " and curdate() between ActiveDate And ifnull(InActiveDate,curdate())"
                Query &= " order by (case when Key2 in ('', '*') then 9 else 1 end)"
                Query &= " limit 1"
                MCom.CommandText = Query

                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@oExpedition", oExpedition)

                Dim MaxInputWeight As Double = 0
                Try
                    MaxInputWeight = Math.Round(CDbl(MCom.ExecuteScalar) / 1000, 0) 'dalam KG
                Catch ex As Exception
                    MaxInputWeight = 0
                End Try

                Dim oExpAddInfo As String = ""
                Query = "Select AddInfo From Account Where Account = @oExpedition"
                Try
                    MCom.CommandText = Query

                    MCom.Parameters.Clear()
                    MCom.Parameters.AddWithValue("@oExpedition", oExpedition)

                    oExpAddInfo = ("" & MCom.ExecuteScalar).ToString.Trim.ToUpper
                Catch ex As Exception
                    oExpAddInfo = ""
                End Try

                If oExpAddInfo.Contains("RNDWGTDCM3=Y") Then
                    oWeight = (Math.Round(CDbl(oWeight), 3, MidpointRounding.AwayFromZero)).ToString
                ElseIf oExpAddInfo.Contains("RNDWGTDCM1=Y") Then
                    oWeight = (Math.Round(CDbl(oWeight), 1, MidpointRounding.AwayFromZero)).ToString
                Else

                End If

                If oWeight > MaxInputWeight Then
                    Result(1) = "Berat max. " & MaxInputWeight & " KG"
                    GoTo Skip
                End If

            End If 'dari If oWeight <> "" And oExpedition <> ""


            If oAWB <> "" And oExpedition <> "" Then

                'validasi nomor resi berulang
                Query = "Select count(*) From OtherExpeditionExpense"
                Query &= " Where Expedition = '" & oExpedition & "'"
                Query &= " And ExpenseType = '" & ExpenseType & "'"
                Query &= " And replace( replace( replace(      AWB     ,'-','') ,'.',''), ' ','')"
                Query &= "   = replace( replace( replace('" & oAWB & "','-','') ,'.',''), ' ','')"
                '2022-06-14 10:11, By Cucun, Jika Informasi nya sama persis boleh lewat
                '2022-06-16 09:39, By Cucun, Tidak disetujui cara seperti dibawah, langsung hubungi pa alex
                'Query &= " And ("
                'Query &= "   Ori <> '" & oOri & "'"
                'Query &= "   Or Dst <> '" & oDst & "'"
                'Query &= "   Or Weight <> '" & oWeight & "'"
                'Query &= " )"
                MCom.CommandText = Query
                If MCom.ExecuteScalar > 0 Then
                    Result(1) = "Sudah pernah ada AWB 3PL (" & oExpedition & "." & oAWB & ")"
                    GoTo Skip
                End If

            End If 'dari If oAWB <> "" And oExpedition <> ""


            Dim nMaxLengthOtherAWB As Integer = 300

            Dim OtherExpeditionListAll As String() = OtherExpeditionList.Split("|")

            For i As Integer = 0 To OtherExpeditionListAll.Length - 1

                Dim ConsNum As String = OtherExpeditionListAll(i).Split(",")(0)
                Dim SjNum As String = OtherExpeditionListAll(i).Split(",")(1)
                Dim OtherExp As String = OtherExpeditionListAll(i).Split(",")(2).Trim

                Dim OtherAWB As String = OtherExpeditionListAll(i).Split(",")(3).Trim
                If OtherAWB = "" Then
                    If TrackNum3plByIpp <> "" Then
                        OtherAWB = TrackNum3plByIpp
                    End If
                End If

                Dim TrackNum As String = ""
                Try
                    TrackNum = OtherExpeditionListAll(i).Split(",")(4)
                Catch ex As Exception
                    TrackNum = ""
                End Try

                Dim sWeight As String = ""
                Try
                    sWeight = OtherExpeditionListAll(i).Split(",")(5)
                Catch ex As Exception
                    sWeight = ""
                End Try

                If sWeight <> "" Then
                    Dim IsWeight As Boolean = True

                    If Not IsNumeric(sWeight) Then
                        IsWeight = False
                        GoTo SkipWeight
                    End If

                    If Math.Round(CDbl(sWeight), 3) <= 0 Then
                        IsWeight = False
                        GoTo SkipWeight
                    End If

SkipWeight:
                    If IsWeight = False Then
                        Result(1) = "Isi berat paket dengan benar"
                        GoTo Skip
                    End If

                End If 'dari If sWeight <> ""

                If OtherExp = "" Then
                    Result(1) = "Isi Ekspedisi / 3PL"
                    GoTo Skip
                End If

                If OtherAWB = "" Then
                    Result(1) = "Isi Nomor Resi 3PL"
                    GoTo Skip
                End If

                If OtherAWB <> "" And OtherAWB.Length > nMaxLengthOtherAWB Then
                    Result(1) = "Isi nomor AWB 3PL dengan benar (max. " & nMaxLengthOtherAWB & " karakter)"
                    GoTo Skip
                End If

                Dim OtherWeight As Double = 0
                Try
                    OtherWeight = CDbl(sWeight)
                Catch ex As Exception
                    OtherWeight = 0
                End Try

                If OtherWeight <= 0 Then
                    If ExpenseType = "PUP" Then
                        'Jemput Rekanan, berdasarkan Jumlah, bukan Berat
                    ElseIf ExpenseType = "FLT" Then
                        'Sewa Armada, berdasarkan Jenis Kendaraan
                    Else
                        Result(1) = "Isi berat paket dengan benar"
                        GoTo Skip
                    End If
                End If

                If OtherWeight > 0 Then
                    Dim OtherExpAddInfo As String = ""
                    Query = "Select AddInfo From Account Where Account = '" & OtherExp & "'"
                    Try
                        MCom.CommandText = Query
                        OtherExpAddInfo = ("" & MCom.ExecuteScalar).ToString.Trim.ToUpper
                    Catch ex As Exception
                        OtherExpAddInfo = ""
                    End Try

                    If OtherExpAddInfo.Contains("RNDWGTDCM3=Y") Then
                        OtherWeight = Math.Round(OtherWeight, 3, MidpointRounding.AwayFromZero)
                    ElseIf OtherExpAddInfo.Contains("RNDWGTDCM1=Y") Then
                        OtherWeight = Math.Round(OtherWeight, 1, MidpointRounding.AwayFromZero)
                    Else
                        OtherWeight = Math.Round(OtherWeight, 3)
                    End If
                End If 'dari If OtherWeight > 0

                Select Case ExpenseType
                    Case "PUP"
                        Query = "Update DispatchE"
                        Query &= " Set PickupExpedition = '" & OtherExp & "', PickupExpeditionAWB = '" & OtherAWB & "', PickupExpeditionAWBDate = '" & oAWBDate & "'"
                        Query &= " , PickupExpeditionNVehicle = '" & oQtyVehicle & "', PickupExpeditionNPickup = '" & oQtyPickup & "'"
                        Query &= " , PickupExpeditionUpdTime = now(), PickupExpeditionUpdUser = '" & UserId & "'"
                        Query &= " Where Station = '" & OriCode & "' And SjNum = '" & SjNum & "'"
                        Try
                            MCom.CommandText = Query
                            MCom.ExecuteNonQuery()
                        Catch ex As Exception
                            Dim e As New ClsError
                            e.ErrorLog(MCon, AppName, AppVersion, Method, OriCode, ex, Query, LogKeyword)

                            Result(1) &= "Error : " & ex.Message
                            GoTo Skip
                        End Try

                    Case "FLT"
                        Query = "Update SuratJalanD"
                        Query &= " Set OtherExpedition = '" & OtherExp & "', OtherExpeditionAWB = '" & OtherAWB & "', OtherExpeditionWeight = '" & OtherWeight & "'"
                        Query &= " , OtherExpeditionOri = '" & oOri & "', OtherExpeditionDst = '" & oDst & "'"
                        Query &= " , OtherExpeditionCost = 0, OtherExpeditionVehicleType = '" & oVehicleType & "'"
                        Query &= " , OtherExpeditionUpdTime = now(), OtherExpeditionUpdUser = '" & UserId & "'"
                        Query &= " Where TrackNum = '" & TrackNum & "' And SjNum = '" & SjNum & "'"
                        Query &= " And `Status` = '0'"

                        Try
                            MCom.CommandText = Query
                            MCom.ExecuteNonQuery()
                        Catch ex As Exception
                            Dim e As New ClsError
                            e.ErrorLog(MCon, AppName, AppVersion, Method, OriCode, ex, Query, LogKeyword)

                            Result(1) &= "Error : " & ex.Message
                            GoTo Skip
                        End Try

                    Case Else
                        If ConsNum <> "" And ConsNum <> "0" And SjNum <> "" Then

                            Query = "Update SuratJalanD"
                            Query &= " Set OtherExpedition = '" & OtherExp & "', OtherExpeditionAWB = '" & OtherAWB & "', OtherExpeditionWeight = '" & OtherWeight & "'"
                            Query &= " , OtherExpeditionOri = '" & oOri & "', OtherExpeditionDst = '" & oDst & "'"
                            Query &= " , OtherExpeditionCost = 0"
                            Query &= " , OtherExpeditionUpdTime = now(), OtherExpeditionUpdUser = '" & UserId & "'"
                            Query &= " Where ConsNum = '" & ConsNum & "' And SjNum = '" & SjNum & "'"
                            Query &= " And `Status` = '0'"
                            Try
                                MCom.CommandText = Query
                                MCom.ExecuteNonQuery()
                            Catch ex As Exception
                                Dim e As New ClsError
                                e.ErrorLog(MCon, AppName, AppVersion, Method, OriCode, ex, Query, LogKeyword)

                                Result(1) &= "Error : " & ex.Message
                                GoTo Skip
                            End Try

                        End If

                        If TrackNum <> "" And SjNum <> "" Then

                            Query = "Update SuratJalanD"
                            Query &= " Set OtherExpedition = '" & OtherExp & "', OtherExpeditionAWB = '" & OtherAWB & "', OtherExpeditionWeight = '" & OtherWeight & "'"
                            Query &= " , OtherExpeditionOri = '" & oOri & "', OtherExpeditionDst = '" & oDst & "'"
                            Query &= " , OtherExpeditionCost = 0"
                            Query &= " , OtherExpeditionUpdTime = now(), OtherExpeditionUpdUser = '" & UserId & "'"
                            Query &= " Where TrackNum = '" & TrackNum & "' And SjNum = '" & SjNum & "'"
                            Query &= " And `Status` = '0'"

                            Try
                                MCom.CommandText = Query
                                MCom.ExecuteNonQuery()
                            Catch ex As Exception
                                Dim e As New ClsError
                                e.ErrorLog(MCon, AppName, AppVersion, Method, OriCode, ex, Query, LogKeyword)

                                Result(1) &= "Error : " & ex.Message
                                GoTo Skip
                            End Try

                        End If
                End Select

            Next


            If oAWB <> "" And oAWB.Length > nMaxLengthOtherAWB Then
                Result(1) = "Isi nomor AWB 3PL dengan benar (max. " & nMaxLengthOtherAWB & " karakter)"
                GoTo Skip
            End If

            If oAWB <> "" And ExpenseType = "CON" Then
                Query = "Select count(Nomor) From ResiKonsol_Data Where @oAWB like concat(Nomor,'%')"
                MCom.Parameters.Clear()
                MCom.Parameters.AddWithValue("@oAWB", oAWB)
                Try
                    MCom.CommandText = Query
                    If MCom.ExecuteScalar > 0 Then
                        Result(1) = "Nomor terdaftar sebagai Resi Konsol"
                        GoTo Skip
                    End If
                Catch ex As Exception
                    Dim e As New ClsError
                    e.ErrorLog(MCon, AppName, AppVersion, Method, OriCode, ex, Query, LogKeyword)

                    Result(1) &= "Error : " & ex.Message
                    GoTo Skip
                End Try
            End If


            If oExpedition <> "" And oAWB <> "" Then

                '2022-06-14 10:15, By Cucun, Pindah keatas supaya oWeight nya seragam antara Validasi dan Insert ke OtherExpeditionExpense
                'Dim oExpAddInfo As String = ""
                'Query = "Select AddInfo From Account Where Account = '" & oExpedition & "'"
                'Try
                '    MCom.CommandText = Query
                '    oExpAddInfo = ("" & MCom.ExecuteScalar).ToString.Trim.ToUpper
                'Catch ex As Exception
                '    oExpAddInfo = ""
                'End Try

                'If oExpAddInfo.Contains("RNDWGTDCM3=Y") Then
                '    oWeight = (Math.Round(CDbl(oWeight), 3, MidpointRounding.AwayFromZero)).ToString
                'ElseIf oExpAddInfo.Contains("RNDWGTDCM1=Y") Then
                '    oWeight = (Math.Round(CDbl(oWeight), 1, MidpointRounding.AwayFromZero)).ToString
                'Else

                'End If

                Dim oParam(15) As Object
                oParam(0) = User
                oParam(1) = Password
                oParam(2) = oExpedition
                oParam(3) = oAWB
                oParam(4) = oOri
                oParam(5) = oDst
                If ExpenseType = "FLT" Then
                    oParam(6) = oQtyVehicle 'Weight di-overwrite dengan Jumlah Kendaraan
                Else
                    oParam(6) = oWeight
                End If
                oParam(7) = oAWBDate
                oParam(8) = "" 'AddInfo
                oParam(9) = oAWBDCost 'bila -1, maka biaya sesuai settingan
                oParam(10) = -1 'CustomWeight
                oParam(11) = ExpenseType
                oParam(12) = oQtyVehicle
                oParam(13) = oQtyPickup
                oParam(14) = oVehicleType
                oParam(15) = UserId

                OtherExpeditionExpenseUpdate(AppName, AppVersion, oParam)

            End If


            Process = True

Skip:

            If Process Then
                MTrn.Commit()

                If IsTrackNum3plByIpp Then
                    ReDim Result(3)
                End If

                Result(0) = "0"
                Result(1) = ResponseOK

                If IsTrackNum3plByIpp Then
                    Result(3) = TrackNum3plByIpp
                End If

                If False Then
                    Try
                        Dim t As New Threading.Thread(AddressOf CostOtherExpeditionAWB)
                        t.Start()
                    Catch ex As Exception
                    End Try
                End If
            Else
                MTrn.Rollback()
            End If

            CreateResponseLog(AppName, AppVersion, Method, User, Result, , LogKeyword)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Private Function ConvertMonthToAlphabet(ByVal Month As Integer) As String

        'Jan = A
        'Feb = B
        'Mar = C
        'Apr = D
        'May = E
        'Jun = F
        'Jul = G
        'Aug = H
        'Sep = I
        'Oct = J
        'Nov = K
        'Dec = L

        Return CStr(Chr(CInt(Month) - 1 + 65)) 'ascii, A=65

    End Function

    Private Sub CostOtherExpeditionAWB()
        CostOtherExpeditionAWBCore()
    End Sub

    <WebMethod()>
    Public Function CostOtherExpeditionAWBCore() As Object()

        Dim Method As String = "CostOtherExpeditionAWB"

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Try
            Dim MTrn As MySqlTransaction

            MCon = MasterMCon.Clone

            Dim MCom As New MySqlCommand("", MCon)
            MCon.Open()

            MTrn = MCon.BeginTransaction()
            MCom.Transaction = MTrn

            Dim Process As Boolean = False

            Query = "Update SuratJalanD sj, SuratJalanH sh, CourierDeliveryCost cost"
            Query &= " Set sj.OtherExpeditionCost = sj.OtherExpeditionWeight * cost.Rate" '-- update ongkir transaksi
            Query &= " Where sj.OtherExpedition <> '' And sj.OtherExpeditionAWB <> '' And sj.OtherExpeditionWeight <> ''" '-- informasi ekspedisi sudah lengkap
            Query &= " And sj.OtherExpeditionCost = 0" '-- ongkir ekspedisi belum lengkap
            Query &= " And sj.ConsNum <> '0' And sj.SjNum <> ''" '-- pengiriman kons (hub-hub)
            Query &= " And sj.`Status` = '0'" '-- masih dalam perjalanan
            Query &= " And sj.SjNum = sh.SjNum And sj.DriverId = sh.DriverId" '-- link suratjalanh dengan suratjaland
            Query &= " And cost.Account = sj.OtherExpedition" '-- link ekspedisi
            Query &= " And cost.OriType = '1' And cost.Ori = replace(sh.Code,'S','H')" '-- link hub asal
            Query &= " And cost.DstType = '1' And cost.Dst = replace(sj.Dst,'S','H')" '-- link hub tujuan
            Query &= " and curdate() between cost.ActiveDate and ifnull(cost.InactiveDate,curdate())"
            'Query &= " and sh.addtime >= date_add(curdate(), interval -14 day)"
            Try
                MCom.CommandText = Query
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog("", "", Method, "", ex, Query)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try


            Query = "Update SuratJalanD sj, SuratJalanH sh, CourierDeliveryCost cost, `Transaction` t, MstPostalCode pc"
            Query &= " Set sj.OtherExpeditionCost = sj.OtherExpeditionWeight * cost.Rate" '-- update ongkir transaksi
            Query &= " Where sj.OtherExpedition <> '' And sj.OtherExpeditionAWB <> '' And sj.OtherExpeditionWeight <> ''" '-- informasi ekspedisi sudah lengkap
            Query &= " And sj.OtherExpeditionCost = 0" '-- ongkir ekspedisi belum lengkap
            Query &= " And sj.TrackNum <> '' And sj.SjNum <> ''" '-- pengiriman alamat (hub-city)
            Query &= " And sj.`Status` = '0'" '-- masih dalam perjalanan
            Query &= " And sj.SjNum = sh.SjNum And sj.DriverId = sh.DriverId" '-- link suratjalanh dengan suratjaland
            Query &= " And sj.TrackNum = t.TrackNum" '-- link dengan transaction
            Query &= " And t.CoPostalCode = pc.Code" '-- link dengan postalcode
            Query &= " And cost.Account = sj.OtherExpedition" '-- link ekspedisi
            Query &= " And cost.OriType = '1' And cost.Ori = replace(sh.Code,'S','H')" '-- link hub asal
            Query &= " And cost.DstType = '2' And cost.Dst = pc.City" '-- link kota tujuan
            Query &= " and curdate() between cost.ActiveDate and ifnull(cost.InactiveDate,curdate())"
            'Query &= " and sh.addtime >= date_add(curdate(), interval -14 day)"
            Try
                MCom.CommandText = Query
                MCom.ExecuteNonQuery()
            Catch ex As Exception
                Dim e As New ClsError
                e.ErrorLog("", "", Method, "", ex, Query)

                Result(1) &= "Error : " & ex.Message
                GoTo Skip
            End Try

            Process = True

Skip:

            If Process Then
                MTrn.Commit()

                Result(0) = "0"
                Result(1) = ResponseOK
            Else
                MTrn.Rollback()
            End If

        Catch ex As Exception
            Dim e As New ClsError
            Dim User As String = ""
            e.ErrorLog("", "", Method, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Private Function ExpeditionExpenseTableName(ByVal ExpenseType As String) As String

        Dim Result As String = ""

        Select Case ExpenseType
            Case "AWB"
                Result = "CourierAddressDeliveryCost"
            Case "PUE"
                Result = "CourierDooringDeliveryCost"
            Case "PUP"
                Result = "CourierPickupDeliveryCost"
            Case "FLT"
                Result = "CourierRentalFleetDeliveryCost"
            Case Else
                'CON
                Result = "CourierDeliveryCost"
        End Select

        Return Result

    End Function

    Private Function GetHub(ByVal MCon As MySqlConnection, ByVal HubCode As String) As String

        Dim Hub As String = ""

        Try
            Dim ObjSQL As New ClsSQL

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@HubCode", HubCode)

            Hub = "" & ObjSQL.ExecScalarWithParam(MCon, "Select `Hub` From MstHub Where `Hub` = @HubCode", SqlParam)

        Catch ex As Exception
            Hub = ""

        End Try

        Return Hub

    End Function

    Private Function GetStation(ByVal MCon As MySqlConnection, ByVal StationCode As String) As String

        Dim Station As String = ""

        Try
            Dim ObjSQL As New ClsSQL

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@StationCode", StationCode)

            Station = "" & ObjSQL.ExecScalarWithParam(MCon, "Select `Station` From MstStation Where `Station` = @StationCode", SqlParam)

        Catch ex As Exception
            Station = ""

        End Try

        Return Station

    End Function

    Private Function GetStationNotFound(ByVal Station As String) As String
        Return "Gateway " & Station & " tidak ditemukan"
    End Function

    Private Function GetStationOrHubNotFound(ByVal Code As String) As String
        Return "Gateway atau Hub " & Code & " tidak ditemukan"
    End Function

    <WebMethod()>
    Public Function OtherExpeditionExpenseUpdate(ByVal AppName As String, ByVal AppVersion As String _
        , ByVal Param As Object()) As Object()

        Dim Method As String = "OtherExpeditionExpenseUpdate"

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString
            Dim Expedition As String = Param(2).ToString.Trim.ToUpper
            Dim AWB As String = Param(3).ToString.Trim.ToUpper
            Dim Ori As String = Param(4).ToString.Trim.ToUpper
            Dim Dst As String = Param(5).ToString.Trim.ToUpper
            Dim Weight As String = Param(6).ToString.Trim.ToUpper
            Dim AWBDate As String = Param(7).ToString.Trim.ToUpper

            Dim AddInfo As String = ""
            Try
                AddInfo = Param(8).ToString.Trim
            Catch ex As Exception
                AddInfo = ""
            End Try

            Dim CustomCost As Double = -1
            Try
                If Not Param(9) Is Nothing Then
                    If IsNumeric(Param(9)) Then
                        CustomCost = CDbl(Param(9))
                        If CustomCost = 0 Then
                            CustomCost = -1
                        End If
                    End If
                End If
            Catch ex As Exception
                CustomCost = -1
            End Try

            Dim CustomWeight As Double = -1
            Try
                CustomWeight = CDbl(Param(10))
            Catch ex As Exception
                CustomWeight = -1
            End Try

            'menentukan biaya ekspedisi yang digunakan
            'CON = Serah Ekspedisi
            'AWB = Serah Antar Alamat
            'PUE = Jemput Ekspedisi
            'PUP = Jemput Partner
            'FLT = Sewa Armada
            Dim ExpenseType As String = ""
            Try
                ExpenseType = Param(11).ToString.Trim.ToUpper
            Catch ex As Exception
                ExpenseType = ""
            End Try
            If ExpenseType = "" Then
                ExpenseType = "CON"
            End If

            Dim ExpenseTbName As String = ExpeditionExpenseTableName(ExpenseType)

            If ExpenseType = "PUE" Then
                Ori = Dst
            End If
            If ExpenseType = "PUP" Then
                Dst = Ori
            End If


            'untuk ExpenseType PUP dan FLT
            Dim QtyVehicle As Integer = 0
            Try
                QtyVehicle = CInt(Param(12))
            Catch ex As Exception
                QtyVehicle = 0
            End Try

            'untuk ExpenseType PUP
            Dim QtyPickup As Integer = 0
            Try
                QtyPickup = CInt(Param(13))
            Catch ex As Exception
                QtyPickup = 0
            End Try


            'untuk ExpenseType FLT
            Dim VehicleType As String = ""
            Try
                VehicleType = Param(14).ToString.Trim.ToUpper
            Catch ex As Exception
                VehicleType = ""
            End Try

            Dim UserId As String = ""
            Try
                UserId = Param(15).ToString.Trim.ToUpper
            Catch ex As Exception
                UserId = ""
            End Try
            If UserId = "" Then
                UserId = User & "_NA"
            End If


            MCon = MasterMCon.Clone

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            If UserOK(0) <> "0" Then
                Result(1) = UserOK(1)
                GoTo Skip
            End If

            Dim LogKeyword As String = Left(Expedition & " " & ExpenseType & " " & AWB, 200)

            CreateRequestLog(AppName, AppVersion, Method, User, "", Param, LogKeyword)

            If Expedition = "" Or AWB = "" Then
                Result(1) = "Input Ekspedisi dan Nomor Resi"
                GoTo Skip
            End If

            If Ori = "" Or Dst = "" Then
                Result(1) = "Input Asal dan Tujuan"
                GoTo Skip
            End If

            If CustomWeight <> -1 Then
                Weight = CustomWeight
            End If

            If Weight = "" Or Weight <= "0" Or IsNumeric(Weight) = False Then
                If ExpenseType = "PUP" Then
                    Weight = "0"
                    Try
                        Weight = (QtyVehicle * QtyPickup)
                    Catch ex As Exception
                        Weight = "0"
                    End Try

                ElseIf ExpenseType = "FLT" Then

                Else
                    Result(1) = "Input Berat KG dengan benar"
                    GoTo Skip
                End If
            End If

            If AWBDate = "" Or IsDate(AWBDate) = False Then
                Result(1) = "Input Tanggal Nomor Resi dengan benar"
                GoTo Skip
            End If

            If ExpenseType = "PUP" Then
                If QtyVehicle < 1 Or QtyPickup < 1 Then
                    Result(1) = "Input Jumlah Kendaraan dan Jumlah Pickup"
                    GoTo Skip
                End If
            End If 'dari If ExpenseType = "PUP"

            If ExpenseType = "FLT" Then
                If VehicleType = "" Then
                    Result(1) = "Input Jenis Kendaraan"
                    GoTo Skip
                End If
                If QtyVehicle < 1 Then
                    Result(1) = "Input Jumlah Kendaraan"
                    GoTo Skip
                End If
            End If 'dari If ExpenseType = "FLT"


            Dim ObjSQL As New ClsSQL


            'Dim AWB3PLBCKDTE As Integer = -1
            'Try
            '    Query = "Select `Value` From Const Where Key1 = 'AWB3PLBCKDTE'"
            '    Dim TempAWB3PLBCKDTE As String = ObjSQL.ExecScalar(MCon, Query).ToString.Trim
            '    If TempAWB3PLBCKDTE = "" Or (IsNumeric(TempAWB3PLBCKDTE) = False) Then
            '        AWB3PLBCKDTE = -1
            '    Else
            '        AWB3PLBCKDTE = CInt(TempAWB3PLBCKDTE)
            '    End If
            'Catch ex As Exception
            '    AWB3PLBCKDTE = -1
            'End Try
            'If AWB3PLBCKDTE = -1 Then
            '    AWB3PLBCKDTE = 3
            'End If

            'Query = "select cast(curdate() as char)"
            'Dim SqlCurDate As String = ObjSQL.ExecScalar(MCon, Query)

            'Query = "Select case when date_add('" & Format(CDate(SqlCurDate), "yyyy-MM-dd") & "', interval -" & AWB3PLBCKDTE & " day) <= '" & Format(CDate(AWBDate), "yyyy-MM-dd") & "' then '1' else '0' end"
            'If ObjSQL.ExecScalar(MCon, Query) < 1 Then
            '    Result(1) = "Tanggal AWB 3PL max. " & AWB3PLBCKDTE & " hari yang lalu (" & Format(DateAdd(DateInterval.Day, -1 * AWB3PLBCKDTE, CDate(SqlCurDate)), "yyyy-MM-dd") & ")"
            '    GoTo Skip
            'End If


            'cek harga 3PL sudah di-setting
            If OtherExpeditionDefault.Contains(Expedition & ".") = False Then
                Query = "Select count(Account) From " & ExpenseTbName
                Query &= " Where Account = '" & Expedition & "'"
                Select Case ExpenseType
                    Case "AWB"
                        Query &= " And Ori = '" & Ori & "' And Dst = '" & Dst & "'" '-- Dst = Kode Kota Tujuan
                    Case "PUE"
                        Query &= " And Dst = '" & Dst & "'" '-- Dst = Kode Hub yang proses
                    Case "PUP"
                        Query &= " And Ori = '" & Ori & "'" '-- Ori = Kode Hub yang proses
                    Case "FLT"
                        Query &= " And Ori = '" & Ori & "' And Dst = '" & Dst & "'" '-- Ori dan Dst = Kode Kota
                    Case Else
                        'CON
                        Query &= " And Ori = '" & Ori & "' And Dst = '" & Dst & "'" '-- Dst = Kode Hub Tujuan
                End Select
                Query &= " and curdate() between ActiveDate and ifnull(InActiveDate,curdate())"
                If ObjSQL.ExecScalar(MCon, Query) < 1 Then
                    If CustomCost > 0 Then
                        'tidak perlu cek Master Biaya Ekspedisi
                    Else
                        Result(1) = "Belum ada harga 3PL (" & Expedition & "." & Ori & "." & Dst & ")"
                        GoTo Skip
                    End If
                End If
            End If


            Query = "SELECT COUNT(*) AS ctr"
            Query &= " FROM OtherExpeditionExpense"
            Query &= " WHERE Expedition = '" & Expedition & "'"
            Query &= " AND AWB = '" & AWB & "'"
            Query &= " AND ExpenseType = '" & ExpenseType & "'"

            Dim AWB3PLExists As Integer = ObjSQL.ExecScalar(MCon, Query)
            If AWB3PLExists > 0 Then

                'Kalau AWB 3PL dan Ekspedisi nya ada di OtherExpeditionExpense
                'cek apakah Ori,Dst,Weight,AWBDate sudah pernah ada
                'Jika ada diberhasilkan, jika berbeda tolak

                Query = "SELECT COUNT(*) AS ctr"
                Query &= " FROM OtherExpeditionExpense"
                Query &= " WHERE Expedition = '" & Expedition & "'"
                Query &= " AND AWB = '" & AWB & "'"
                Query &= " AND ExpenseType = '" & ExpenseType & "'"
                Query &= " AND Ori = '" & Ori & "'"
                Query &= " AND Dst = '" & Dst & "'"
                If ExpenseType = "PUP" Or ExpenseType = "FLT" Then
                    'tidak menggunakan Berat
                Else
                    Query &= " AND Weight = '" & Weight & "'"
                End If
                Query &= " AND AWBDate = '" & AWBDate & "'"

                Dim AWB3PLIdentic As Integer = ObjSQL.ExecScalar(MCon, Query)

                If AWB3PLIdentic > 0 Then
                    Result(0) = "0"
                    Result(1) = ResponseOK
                    GoTo Skip
                Else
                    Result(1) = "AWB 3PL " & AWB & " untuk Ekpedisi " & Expedition & " sudah pernah digunakan!"
                    GoTo Skip
                End If

            End If

            Query = "Insert Into OtherExpeditionExpense ("
            Query &= " Expedition, AWB, ExpenseType"
            Query &= " , OriType, Ori, OriName, DstType, Dst, DstName"
            Query &= " , Weight, AWBDate, AddInfo"
            Query &= " , Cost, Rate, LeadTime, MinWeight, MinRate"
            Query &= " , AddTime, AddUser"
            Query &= " ) values ("
            Query &= " '" & Expedition & "', '" & AWB & "', '" & ExpenseType & "'"
            Select Case ExpenseType
                Case "AWB"
                    Query &= " , '1', '" & Ori & "', '', '2', '" & Dst & "', ''"
                Case "FLT"
                    Query &= " , '2', '" & Ori & "', '', '2', '" & Dst & "', ''"
                Case Else
                    Query &= " , '1', '" & Ori & "', '', '1', '" & Dst & "', ''"
            End Select
            Query &= " , '" & Weight & "', '" & AWBDate & "', '" & AddInfo & "'"
            Query &= " , 0, 0, '', 0, 0"
            Query &= " , now(), '" & UserId & "'"
            Query &= " )"
            '2022-04-05 16:59, By Cucun, Untuk kedepannya data OtherExpeditionExpense ga boleh gerak2
            'Dikomplain sama finance acrued mingguan dan akhir bulan nilai nya berbeda
            'Query &= " ) on duplicate key update"
            'Query &= " Ori = '" & Ori & "', Dst = '" & Dst & "', Weight = '" & Weight & "', AWBDate = '" & AWBDate & "', AddInfo = '" & AddInfo & "'"
            'Query &= " , Cost = 0, Rate = 0, LeadTime = '', MinWeight = 0, MinRate = 0"
            'Query &= " , UpdTime = now(), UpdUser = '" & User & "'"
            If ObjSQL.ExecNonQuery(MCon, Query) = False Then
                Result(1) = "Gagal update data"
                GoTo Skip
            End If

            'update nama asal
            Select Case ExpenseType
                Case "FLT"
                    'asal pakai kode kota
                    Query = "Update OtherExpeditionExpense o, MstCity h"
                    Query &= " Set o.OriName = h.Name"
                    Query &= " Where o.Ori = h.City"
                    Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                    Query &= " And o.ExpenseType = '" & ExpenseType & "'"
                Case Else
                    'asal pakai kode hub
                    Query = "Update OtherExpeditionExpense o, MstHub h"
                    Query &= " Set o.OriName = h.Alias"
                    Query &= " Where o.Ori = h.Hub"
                    Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                    Query &= " And o.ExpenseType = '" & ExpenseType & "'"
            End Select
            If ObjSQL.ExecNonQuery(MCon, Query) = False Then
                Result(1) = "Gagal update Asal"
                GoTo Skip
            End If

            'update nama tujuan
            Select Case ExpenseType
                Case "AWB", "FLT"
                    'tujuan pakai kode kota
                    Query = "Update OtherExpeditionExpense o, MstCity h"
                    Query &= " Set o.DstName = h.Name"
                    Query &= " Where o.Dst = h.City"
                    Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                    Query &= " And o.ExpenseType = '" & ExpenseType & "'"
                Case Else
                    'tujuan pakai kode hub
                    Query = "Update OtherExpeditionExpense o, MstHub h"
                    Query &= " Set o.DstName = h.Alias"
                    Query &= " Where o.Dst = h.Hub"
                    Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                    Query &= " And o.ExpenseType = '" & ExpenseType & "'"
            End Select
            If ObjSQL.ExecNonQuery(MCon, Query) = False Then
                Result(1) = "Gagal update Tujuan"
                GoTo Skip
            End If

            If CustomCost = -1 Then

                'antisipasi table Biaya Ekspedisi bisa multi ActiveDate - 2023-09-15
                'catat data Biaya Ekspedisi yang aktif terakhir sebelum update ke OtherExpeditionExpense
                Dim ExpenseTbNameActive As String = "Temp" & ExpenseTbName & "Active"

                Query = "Drop Temporary Table If Exists " & ExpenseTbNameActive
                ObjSQL.ExecNonQueryWithParam(MCon, Query, Nothing)

                Query = "Create Temporary Table " & ExpenseTbNameActive & " Like " & ExpenseTbName
                ObjSQL.ExecNonQueryWithParam(MCon, Query, Nothing)

                Select Case ExpenseType
                    Case "PUP"

                        Query = "Insert Into " & ExpenseTbNameActive
                        Query &= " Select c.*"
                        Query &= " From OtherExpeditionExpense o, " & ExpenseTbName & " c"
                        Query &= " Where o.Expedition = c.Account"
                        Query &= " And o.OriType = c.OriType and o.Ori = c.Ori"
                        Query &= " And curdate() between c.ActiveDate and ifnull(c.InActiveDate,curdate())"
                        Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                        Query &= " And o.ExpenseType = '" & ExpenseType & "'"
                        Query &= " Order By c.ActiveDate Desc limit 1"
                        ObjSQL.ExecNonQueryWithParam(MCon, Query, Nothing)

                        'Query = "Update OtherExpeditionExpense o, " & ExpenseTbName & " c"
                        Query = "Update OtherExpeditionExpense o, " & ExpenseTbNameActive & " c"
                        Query &= " Set o.Rate = c.Rate, o.MinRate = c.Rate"
                        Query &= " , o.LeadTime = '1', o.MinWeight = '1'"
                        Query &= " , o.Weight = " & Weight '-- Weight = (QtyVehicle * QtyPickup)
                        Query &= " , o.Cost      = round(" & Weight & " * c.Rate, 0)"
                        Query &= " , o.CostDraft = round(" & Weight & " * c.Rate, 0)"
                        'Ongkir Akhir = Jumlah Kendaraan * Jumlah Pickup * Ongkir Jemput

                        Query &= " Where o.Expedition = c.Account"
                        Query &= " And o.OriType = c.OriType and o.Ori = c.Ori"
                        Query &= " And curdate() between c.ActiveDate and ifnull(c.InActiveDate,curdate())"
                        Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                        Query &= " And o.ExpenseType = '" & ExpenseType & "'"

                    Case "FLT"

                        Query = "Insert Into " & ExpenseTbNameActive
                        Query &= " Select c.*"
                        Query &= " From OtherExpeditionExpense o, " & ExpenseTbName & " c"
                        Query &= " Where o.Expedition = c.Account"
                        Query &= " And o.Ori = c.Ori And o.Dst = c.Dst "
                        Query &= " And curdate() between c.ActiveDate and ifnull(c.InActiveDate,curdate())"
                        Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                        Query &= " And o.ExpenseType = '" & ExpenseType & "' And c.VehicleType = '" & VehicleType & "'"
                        Query &= " Order By c.ActiveDate Desc limit 1"
                        ObjSQL.ExecNonQueryWithParam(MCon, Query, Nothing)

                        'Query = "Update OtherExpeditionExpense o, " & ExpenseTbName & " c"
                        Query = "Update OtherExpeditionExpense o, " & ExpenseTbNameActive & " c"
                        Query &= " Set o.Rate = c.Rate, o.MinRate = c.Rate"
                        Query &= " , o.LeadTime = c.LeadTime, o.MinWeight = '1'"
                        Query &= " , o.Cost = round(o.Weight * c.Rate, 0), o.CostDraft = round(o.Weight * c.Rate, 0)" '-- Weight = Jumlah Kendaraan Sewa Armada
                        'Ongkir Akhir = Jumlah Kendaraan Sewa Armada * Ongkir Jemput

                        Query &= " Where o.Expedition = c.Account"
                        Query &= " And o.Ori = c.Ori And o.Dst = c.Dst "
                        Query &= " And curdate() between c.ActiveDate and ifnull(c.InActiveDate,curdate())"
                        Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                        Query &= " And o.ExpenseType = '" & ExpenseType & "' And c.VehicleType = '" & VehicleType & "'"

                    Case Else

                        Query = "Insert Into " & ExpenseTbNameActive
                        Query &= " Select c.*"
                        Query &= " From OtherExpeditionExpense o, " & ExpenseTbName & " c"
                        Query &= " Where o.Expedition = c.Account"
                        Select Case ExpenseType
                            Case "AWB"
                                Query &= " And o.OriType = c.OriType and o.Ori = c.Ori"
                                Query &= " And o.DstType = c.DstType and o.Dst = c.Dst"
                            Case "PUE"
                                Query &= " And o.OriType = c.DstType and o.Ori = c.Dst"
                                Query &= " And o.DstType = c.DstType and o.Dst = c.Dst"
                            Case Else
                                'CON
                                Query &= " And o.OriType = c.OriType and o.Ori = c.Ori"
                                Query &= " And o.DstType = c.DstType and o.Dst = c.Dst"
                        End Select
                        Query &= " And curdate() between c.ActiveDate and ifnull(c.InActiveDate,curdate())"
                        Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                        Query &= " And o.ExpenseType = '" & ExpenseType & "'"
                        Query &= " Order By c.ActiveDate Desc limit 1"
                        ObjSQL.ExecNonQueryWithParam(MCon, Query, Nothing)

                        'update biaya berdasarkan settingan
                        'Query = "Update OtherExpeditionExpense o, " & ExpenseTbName & " c"
                        Query = "Update OtherExpeditionExpense o, " & ExpenseTbNameActive & " c"
                        Query &= " Set o.Rate = c.Rate, o.MinRate = c.MinRate"
                        Query &= " , o.LeadTime = c.LeadTime, o.MinWeight = c.MinWeight"

                        'Query &= " , o.Cost = o.Weight * c.Rate"
                        Query &= " , o.Cost      = c.MinRate + (case when (o.Weight - c.MinWeight) > 0 then (o.Weight - c.MinWeight) * c.Rate else 0 end)"
                        Query &= " , o.CostDraft = c.MinRate + (case when (o.Weight - c.MinWeight) > 0 then (o.Weight - c.MinWeight) * c.Rate else 0 end)"
                        'Ongkir Akhir = Ongkir Minimum + ((Berat Aktual - Berat Min) * Ongkir Kiloan)

                        Query &= " Where o.Expedition = c.Account"
                        Select Case ExpenseType
                            Case "AWB"
                                Query &= " And o.OriType = c.OriType and o.Ori = c.Ori"
                                Query &= " And o.DstType = c.DstType and o.Dst = c.Dst"
                            Case "PUE"
                                Query &= " And o.OriType = c.DstType and o.Ori = c.Dst"
                                Query &= " And o.DstType = c.DstType and o.Dst = c.Dst"
                            Case Else
                                'CON
                                Query &= " And o.OriType = c.OriType and o.Ori = c.Ori"
                                Query &= " And o.DstType = c.DstType and o.Dst = c.Dst"
                        End Select
                        Query &= " And curdate() between c.ActiveDate and ifnull(c.InActiveDate,curdate())"
                        Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                        Query &= " And o.ExpenseType = '" & ExpenseType & "'"

                End Select

            Else

                'update biaya berdasarkan inputan user
                Query = "Update OtherExpeditionExpense o"
                Query &= " Set o.Rate = '" & CustomCost & "', o.MinRate = '" & CustomCost & "'"
                Query &= " , o.LeadTime = '1', o.MinWeight = '1'"
                Query &= " , o.Cost = '" & CustomCost & "', o.CostDraft = '" & CustomCost & "'"
                Query &= " Where 1=1"
                Query &= " And o.Expedition = '" & Expedition & "' and o.AWB = '" & AWB & "'"
                Query &= " And o.ExpenseType = '" & ExpenseType & "'"

            End If

            'If ObjSQL.ExecNonQuery(MCon, Query) = False Then
            '    Result(1) = "Gagal update Biaya"
            '    GoTo Skip
            'End If
            'tidak perlu error kalau tidak ada perubahan data, tapi catat log
            If ObjSQL.ExecNonQuery(MCon, Query) Then
            Else
                Dim e As New ClsError
                e.DebugLog(MCon, AppName, AppVersion, Method, User, "Gagal OtherExpeditionExpenseUpdate " & Query, LogKeyword)
            End If

            Result(0) = "0"
            Result(1) = ResponseOK

Skip:

            CreateResponseLog(AppName, AppVersion, Method, User, Result, , LogKeyword)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    <WebMethod()>
    Public Function GetHubProcessTypeExpedition(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object(), ByRef dsData As DataSet) As Object()

        Dim Method As String = "GetHubProcessTypeExpedition"

        Dim Result = CreateResult()

        Dim MCon As New MySqlConnection
        Dim MCon2 As New MySqlConnection

        Dim Query As String = ""

        Dim User As String = Param(0).ToString
        Try
            Dim Password As String = Param(1).ToString
            Dim HubCode As String = Param(2).ToString.Trim.ToUpper
            Dim ProcessType As String = Param(3).ToString.Trim.ToUpper

            MCon = MasterMCon.Clone
            MCon2 = Master2MCon.Clone

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            'Result(1) = UserOK(1)
            'GoTo Skip
            'End If

            Dim ObjSql As New ClsSQL

            Query = "Select count(Hub) From MstHub Where Hub = @HubCode"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@HubCode", HubCode)

            If ObjSql.ExecScalarWithParam(MCon2, Query, SqlParam) < 1 Then
                Result(1) = "Hub " & HubCode & " tidak ditemukan"
                GoTo Skip
            End If

            Select Case ProcessType
                Case "SERAH", "TERIMA", "JEMPUT REKANAN", "JEMPUT EKSPEDISI", "SEWA ARMADA", "LASTMILEOUT"

                Case ""
                    Result(1) = "Tipe Proses wajib diisi"
                    GoTo Skip

                Case Else
                    Result(1) = "Proses " & ProcessType & " tidak dikenali"
                    GoTo Skip

            End Select


            Query = "Select a.Account, a.Alias, cast((case when e.Expedition is null then '0' else '1' end) as char) as `Status`"
            Query &= " From Account a"
            Query &= " Left Join MstHubExpedition e on ( e.Expedition = a.Account"
            Query &= " And e.Hub = @HubCode And e.ProcessType = @ProcessType )"
            Query &= " Where curdate() between a.ActiveDate and ifnull(a.InActiveDate,curdate())"
            Query &= " And a.`Type` = '3'"
            Query &= " Order by a.Alias"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@HubCode", HubCode)
            SqlParam.Add("@ProcessType", ProcessType)

            Dim dtQuery As DataTable = ObjSql.ExecDatatableWithParam(MCon2, Query, SqlParam)

            If dtQuery Is Nothing Then
                Result(1) = "Gagal Query"
                GoTo Skip
            End If

            dsData = New DataSet
            dtQuery.TableName = "DATA"
            dsData.Tables.Add(dtQuery)

            Result(0) = "0"
            Result(1) = ResponseOK
            Result(2) = ""

Skip:

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

        Finally
            Try
                MCon.Close()
            Catch ex As Exception
            End Try
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            Try
                MCon2.Close()
            Catch ex As Exception
            End Try
            Try
                MCon2.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result

    End Function

    <WebMethod()>
    Public Function NewExpeditionDownloadTemplate(ByVal AppName As String, ByVal AppVersion As String _
    , ByVal Param As Object(), ByRef dsTemplate As DataSet) As Object()

        Dim Method As String = "NewExpeditionDownloadTemplate"

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            MCon = MasterMCon.Clone
            MCon.Open()

            Dim Process As Boolean = False

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            '    Result(1) = UserOK(1)
            '    GoTo Skip
            'End If

            CreateRequestLog(AppName, AppVersion, Method, User, "", Param)

            Dim ObjSQL As New ClsSQL

            dsTemplate = New DataSet


            Dim dtAccount As New DataTable
            dtAccount.Columns.Add("No")
            dtAccount.Columns.Add("Tipe")
            dtAccount.Columns.Add("Ekspedisi")
            dtAccount.Columns.Add("Nama")
            dtAccount.Columns.Add("Ori")
            dtAccount.Columns.Add("Dest")
            dtAccount.Columns.Add("Resi")
            dtAccount.Columns.Add("Tanggal")
            dtAccount.Columns.Add("Periode Invoice")
            dtAccount.Columns.Add("Keterangan BA")

            'dtAccount.Rows.Add()

            dtAccount.TableName = "ACCOUNT"
            dsTemplate.Tables.Add(dtAccount)

            Process = True

Skip:

            If Process Then

                Result(1) = ResponseOK
                Result(0) = "0"
                Result(2) = ""

            Else

                dsTemplate = Nothing

            End If

            CreateResponseLog(AppName, AppVersion, Method, User, Result)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Public Function GetPackageList(ByVal AppName As String, ByVal AppVersion As String, ByVal Param As Object()) As Object()

        Dim Method As String = "GetPackageList"

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            MCon = MasterMCon.Clone

            'Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            'If UserOK(0) <> "0" Then
            '    Result(1) = UserOK(1)
            '    GoTo Skip
            'End If

            'CreateRequestLog(AppName, AppVersion, Method, User, "", Param)

            Dim ObjSQL As New ClsSQL

            Query = "Select mpc.Category as Kategori, mpc.Description as Deskripsi"
            Query &= " From MstPackageCategory mpc"

            Dim dtQuery As DataTable = ObjSQL.ExecDatatable(MCon, Query)

            If dtQuery Is Nothing Then
                Result(1) = "Gagal query"
                GoTo Skip
            End If

            If dtQuery.Rows.Count < 1 Then
                Result(1) = "Tidak ditemukan"
                GoTo Skip
            End If

            If Not dtQuery Is Nothing Then

                Result(2) = ConvertDatatableToJSON(dtQuery)
                Result(1) = ResponseOK
                Result(0) = "0"

            End If

Skip:

            'CreateResponseLog(AppName, AppVersion, Method, User, Result)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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

    Private Function ConvertDatatableToJSON(ByVal dtResult As DataTable) As String

        On Error Resume Next

        Dim Result As String = ""

        Dim ObjJSON As ClsJSON()
        ReDim ObjJSON(dtResult.Rows.Count - 1)

        For i As Integer = 0 To dtResult.Rows.Count - 1
            ObjJSON(i) = New ClsJSON

            For j As Integer = 0 To dtResult.Columns.Count - 1
                If j = 0 Then
                    ObjJSON(i).Column0 = dtResult.Rows(i).Item(j).ToString
                ElseIf j = 1 Then
                    ObjJSON(i).Column1 = dtResult.Rows(i).Item(j).ToString
                ElseIf j = 2 Then
                    ObjJSON(i).Column2 = dtResult.Rows(i).Item(j).ToString
                ElseIf j = 3 Then
                    ObjJSON(i).Column3 = dtResult.Rows(i).Item(j).ToString
                ElseIf j = 4 Then
                    ObjJSON(i).Column4 = dtResult.Rows(i).Item(j).ToString
                ElseIf j = 5 Then
                    ObjJSON(i).Column5 = dtResult.Rows(i).Item(j).ToString
                End If
            Next
        Next

        Result = JsonConvert.SerializeObject(ObjJSON)
        Result = Replace(Result, "\\", "\")

        'sesuaikan dengan jumlah kolom di ClsJSON
        For j As Integer = 0 To dtResult.Columns.Count - 1
            If j = 0 Then
                Result = Result.Replace("Column0", dtResult.Columns(j).ColumnName.ToUpper)
            ElseIf j = 1 Then
                Result = Result.Replace("Column1", dtResult.Columns(j).ColumnName.ToUpper)
            ElseIf j = 2 Then
                Result = Result.Replace("Column2", dtResult.Columns(j).ColumnName.ToUpper)
            ElseIf j = 3 Then
                Result = Result.Replace("Column3", dtResult.Columns(j).ColumnName.ToUpper)
            ElseIf j = 4 Then
                Result = Result.Replace("Column4", dtResult.Columns(j).ColumnName.ToUpper)
            ElseIf j = 5 Then
                Result = Result.Replace("Column5", dtResult.Columns(j).ColumnName.ToUpper)
            End If
        Next

        'sesuaikan dengan jumlah kolom di ClsJSON
        For j As Integer = 5 To 0 Step -1
            If j = 0 Then
                Result = Result.Replace(",""Column0"":""""", "")

            ElseIf j = 1 Then
                If Result.Contains("""Column1"":") Then
                    Result = Result.Replace(",""Column1"":""""", "")
                Else
                    Exit For
                End If

            ElseIf j = 2 Then
                If Result.Contains("""Column2"":") Then
                    Result = Result.Replace(",""Column2"":""""", "")
                Else
                    Exit For
                End If

            ElseIf j = 3 Then
                If Result.Contains("""Column3"":") Then
                    Result = Result.Replace(",""Column3"":""""", "")
                Else
                    Exit For
                End If

            ElseIf j = 4 Then
                If Result.Contains("""Column4"":") Then
                    Result = Result.Replace(",""Column4"":""""", "")
                Else
                    Exit For
                End If

            ElseIf j = 5 Then
                If Result.Contains("""Column5"":") Then
                    Result = Result.Replace(",""Column5"":""""", "")
                Else
                    Exit For
                End If

            End If
        Next

        Return Result

    End Function

    <WebMethod()>
    Public Function DraftRAOListAll(ByVal AppName As String, ByVal AppVersion As String _
    , ByVal Hub As String, ByVal Param As Object()) As Object()

        Hub = Hub.ToUpper

        Dim Result As Object() = CreateResult()

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString

        Try
            Dim Password As String = Param(1).ToString

            MCon = MasterMCon.Clone

            Dim UserOK As Object() = ValidasiUser(MCon, User, Password)
            If UserOK(0) <> "0" Then
                Result(1) = UserOK(1)
            Else

                CreateRequestLog(AppName, AppVersion, "DraftRAOListAll", User, Hub, Param)

                Dim ObjSQL As New ClsSQL

                'tampilkan daftar draft RAO
                Query = "Select o.PickupNo"
                Query &= " , cast(date_format(date(max(d.UpdTime)),'%Y-%m-%d') as char(10)) as 'RequestTime'"
                Query &= " , count(o.TrackNum) as Paket"
                Query &= " , min(o.`Status`) as `Status`"
                Query &= " From RAODraft o, DCRequest d"
                Query &= " Where o.Hub = '" & Hub & "'"
                Query &= " And o.Hub = d.Hub And o.PickupNo = d.PickupNo"
                Query &= " Group By o.PickupNo"
                Query &= " Order By RequestTime desc"

                Dim dtTemp As DataTable = ObjSQL.ExecDatatable(MCon, Query)

                If Not dtTemp Is Nothing Then

                    Try

                        'tambahkan dengan yang RAI, tapi belum ada request picking dari DC
                        Query = "Select '-' as PickupNo"
                        Query &= " , cast(date_format(date('2000-12-31 00:00:00'),'%Y-%m-%d') as char(10)) as RequestTime"
                        Query &= " , Count(i.TrackNum) as Paket"
                        Query &= " , '9' as `Status`"
                        Query &= " From RackIn i"
                        Query &= " Where i.Hub = '" & Hub & "'"
                        Query &= " And i.TrackNum not in ("
                        Query &= " Select o.TrackNum From RAODraft o"
                        Query &= " Where i.Hub = o.Hub And i.TrackNum = o.TrackNum"
                        Query &= " )"
                        Dim dtTemp2 As DataTable = ObjSQL.ExecDatatable(MCon, Query)

                        'penggabungan daftar
                        If Not dtTemp2 Is Nothing Then
                            If dtTemp2.Rows(0).Item("Paket") <> "0" Then
                                dtTemp.ImportRow(dtTemp2.Rows(0))
                            End If
                        End If

                    Catch ex As Exception
                        Dim e As New ClsError
                        e.ErrorLog(AppName, AppVersion, "DraftRAOListAll", User, ex, Query)

                    End Try

                    Result(2) = ConvertDatatableToString(dtTemp)
                    Result(0) = "0"
                    Result(1) = ResponseOK

                End If

            End If

            CreateResponseLog(AppName, AppVersion, "DraftRAOListAll", User, Result)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, "DraftRAOListAll", User, ex, Query)

            Result(1) &= "Error : " & ex.Message

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



End Class
