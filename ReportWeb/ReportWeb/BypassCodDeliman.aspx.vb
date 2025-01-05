Imports System.Data
Imports System.IO
Imports MySql.Data.MySqlClient
Imports ReportWeb.ClsWebVer
Imports System
Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.Text
Imports Org.BouncyCastle.Crypto.Engines

Partial Class BypassCodDeliman
    Inherits System.Web.UI.Page

    Dim MyPage As String
    Dim ObjFungsi As New ClsFungsi
    Dim ObjCon As New ClsConnection
    Dim ObjSQL As New ClsSQL
    Dim ObjService As New ClsService
    Dim serv As New CoreService.Service
    Dim param() As Object
    Dim respon() As Object
    Dim ScriptManager1 As New ScriptManager

    Dim PageTimeout As Integer = 3600000 'In miliseconds

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("NIK") = "2015548000"
        Session("NameOfUser") = "Ryu"

        ScriptManager1 = ScriptManager.GetCurrent(Me.Page)
        'PageTimeout in miliseconds
        ScriptManager1.AsyncPostBackTimeout = PageTimeout / 1000 'In Seconds

        MyPage = System.Web.VirtualPathUtility.GetFileName(Request.RawUrl).Replace(".aspx", "")
        Dim MyLi As String = "li" & MyPage

        'Dim AllowedMenu As String = "" & Session("AllowedMenu")
        'If AllowedMenu.Contains(MyLi) Then

        Dim MyLink As String = "Link" & MyPage
        Session("CurrentMenu") = MyLink

        If Not IsPostBack Then
            LoadData()
            If Not IsNothing(Session("ResultMessage" & MyPage)) Then
                lblError.Text = Session("ResultMessage" & MyPage)
                Session("ResultMessage" & MyPage) = Nothing
            End If

        End If
        'Else
        'Response.Redirect("notauthorized.aspx", False)
        'End If

    End Sub

    Private Sub LoadData()

        Try
            Dim dt As New DataTable

            dt.TableName = ""
            dt.Columns.Add("Display")
            dt.Columns.Add("Value")

            dt.Rows.Add("", "")
            dt.Rows.Add("KLIKIDM", "KLIK")
            dt.Rows.Add("APKA", "APKA")
            dt.Rows.Add("SPI", "SPI")
            dt.Rows.Add("KLIKIGR", "KIGR")
            ObjFungsi.ddlBound(ddlPaymentType, dt, "Display", "Value")

            dt = New DataTable
            dt.TableName = ""
            dt.Columns.Add("Value")
            dt.Columns.Add("Display")

            dt.Rows.Add("", "")
            dt.Rows.Add("Ganti Alat Bayar", "Ganti Alat Bayar")
            dt.Rows.Add("Pesanan dibatalkan di IGR", "Pesanan dibatalkan di IGR")
            ObjFungsi.ddlBound(ddlReason, dt, "Display", "Value")

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message
        End Try
    End Sub

    Private Function validasiCari() As Boolean

        Try
            lblError.Text = ""

            If ddlPaymentType.SelectedValue = "" Then
                lblError.Text = "Pilih Tipe !"
                Return False
            End If

            TxtPaymentCode.Text = TxtPaymentCode.Text.Trim
            If TxtPaymentCode.Text = "" Then
                lblError.Text = "Isi Nomor Order / Kode Bayar !"
                Return False
            End If

            Return True

        Catch ex As Exception

            lblError.Text = ex.Message
            Return False

        End Try

    End Function

    Protected Sub btnCari_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            If validasiCari() Then

                Dim ErrorMessage As String = ""

                Dim dt As DataTable = bypass_cod_deliman_get(ddlPaymentType.SelectedValue, TxtPaymentCode.Text, ErrorMessage)

                If Not IsNothing(dt) Then

                    If dt.Rows.Count > 0 Then

                        Dim Result0 As String = dt.Rows(0).Item("Result0")
                        Dim Result1 As String = dt.Rows(0).Item("Result1")

                        If Result0 = "0" Then

                            ViewState("gvData") = dt

                            gvData.DataSource = dt
                            gvData.DataBind()

                            trData.Visible = True

                        Else

                            lblError.Text = Result1

                        End If

                    Else

                        lblError.Text = "Tidak ada data!"

                    End If

                Else

                    lblError.Text = ErrorMessage

                End If

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Protected Sub gvData_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs)

        If Not IsNothing(Session("NIK")) Then

            Try
                Dim ErrorMessage As String = ""

                'Dim UpdateType As String = "" & e.CommandName

                'Dim UpdateTypeUpper As String = UpdateType.ToUpper

                Dim IndexRow As Integer = Convert.ToInt32(e.CommandArgument)

                Dim dt As DataTable = ViewState("gvData")

                Dim Reason As String = ddlReason.SelectedValue

                If Reason = "" Then
                    lblError.Text = "Alasan tidak boleh kosong!"
                    GoTo Skip
                End If

                If Not IsNothing(dt) Then
                    If dt.Rows.Count > 0 Then

                        Dim PaymentType As String = dt.Rows(IndexRow).Item("CodPaymentBiller")
                        Dim PaymentCode As String = dt.Rows(IndexRow).Item("CodPaymentCode")

                        Dim dtUpdate As DataTable = bypass_cod_deliman_insert(PaymentType, PaymentCode, Reason, ErrorMessage)

                        If Not IsNothing(dtUpdate) Then
                            If dtUpdate.Rows.Count > 0 Then

                                Dim ResulValue As String = "" & dtUpdate.Rows(0).Item("ResultValue")

                                If ResulValue.StartsWith("ERROR-") Then

                                    Dim Keterangan As String = ResulValue.Substring(6, ResulValue.Length - 6)

                                    lblError.Text = "ByPass Kode Bayar COD gagal! - " & Keterangan

                                Else

                                    lblError.Text = "Berhasil!"

                                    gvData.DataSource = Nothing
                                    gvData.DataBind()

                                    ddlPaymentType.SelectedValue = ""
                                    TxtPaymentCode.Text = ""
                                    TrReason.Visible = False

                                End If

                            End If
                        End If

                    End If
                End If
Skip:
            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
                ObjFungsi.WriteTracelogTxt(Pesan)

            End Try

        End If

    End Sub

    Private Function bypass_cod_deliman_get(ByVal PaymentType As String, ByVal PaymentCode As String, ByRef ErrorMessage As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim SqlQuery As String = ""
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = "call `bypass_cod_deliman_get`( @PaymentType, @PaymentCode )"

            Dim SqlParam As New Dictionary(Of String, String)
            SqlParam.Add("@PaymentType", PaymentType)
            SqlParam.Add("@PaymentCode", PaymentCode)

            MCon = ObjCon.SetConn_Slave2
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam, ErrorMessage)

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message
            lblError.Text = ex.Message

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

    Private Function bypass_cod_deliman_insert(ByVal PaymentType As String, ByVal PaymentCode As String, ByVal Reason As String, ByRef ErrorMessage As String) As DataTable

        Dim dt As DataTable = Nothing

        If Not IsNothing(Session("NIK")) Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing

            Try

                SqlQuery = "call `bypass_cod_deliman_insert`( @PaymentType, @PaymentCode, @Reason, @NIK )"

                Dim SqlParam As New Dictionary(Of String, String)
                SqlParam.Add("@PaymentType", PaymentType)
                SqlParam.Add("@PaymentCode", PaymentCode)
                SqlParam.Add("@Reason", Reason)
                SqlParam.Add("@NIK", "" & Session("NIK"))


                MCon = ObjCon.SetConn_Master
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam, ErrorMessage)

            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
                ObjFungsi.WriteTracelogTxt(Pesan)

                ErrorMessage = ex.Message
                lblError.Text = ex.Message

            Finally

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

            End Try

        End If

        Return dt

    End Function

    Protected Sub gvData_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs)

        'Do Nothing
        '2022-04-01 15:34, By Cucun, e.CommandName = Cancel, mentrigger error The GridView 'gvData' fired event RowCancelingEdit which wasn't handled.

    End Sub

End Class
