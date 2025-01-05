
Imports System.Data
Imports ReportWeb.ClsWebVer
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Collections.Generic

Partial Class BypassPinDeliman
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
        'Response.Redirect("NotAuthorized.aspx", False)
        'End If

    End Sub

    Private Sub LoadData()

        Try
            Dim ErrorList As New StringBuilder
            Dim dt As New DataTable

            dt = New DataTable
            dt.Columns.Add("Value")
            dt.Columns.Add("Display")

            dt.Rows.Add(New String() {"", ""})
            dt.Rows.Add(New String() {"Ganti Alat Bayar", "Ganti Alat Bayar"})
            dt.Rows.Add(New String() {"Pesanan dibatalkan di IGR", "Pesanan dibatalkan di IGR"})
            ObjFungsi.ddlBound(ddlReason, dt, "Display", "Value")

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Function validasiCari(ByRef NoAWB As String) As Boolean

        Try

            lblError.Text = ""

            ViewState("gvData") = Nothing
            gvData.DataSource = Nothing
            gvData.DataBind()

            trData.Visible = False

            NoAWB = TxtAwb.Text.Trim.Replace(vbCrLf, "|")

            Dim AWB() As String = NoAWB.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)

            If AWB.Length > 20 Then
                lblError.Text = "Maksimal Jumlah AWB 20 !!"
                ScriptManager1.SetFocus(TxtError)
                Return False
            End If

            ViewState("BypassReason") = ddlReason.SelectedValue

            Return True

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

            Return False

        End Try

    End Function

    Protected Sub btnCari_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try

            Dim NoAWB As String = ""

            If validasiCari(NoAWB) Then

                Dim ErrorMessage As String = ""

                Dim dt As DataTable = bypassdelimanpin_get(NoAWB, ErrorMessage)


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
                            SetFocus(TxtError)

                        End If

                    Else

                        lblError.Text = "Tidak ada data!"
                        SetFocus(TxtError)

                    End If

                Else

                    lblError.Text = ErrorMessage
                    SetFocus(TxtError)

                End If

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message
            ScriptManager1.SetFocus(TxtError)

        End Try

    End Sub

    Protected Sub gvData_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs)

        If Not IsNothing(Session("NIK")) Then

            Try

                Dim ErrorMessage As String = ""

                Dim UpdateType As String = "" & e.CommandName

                Dim UpdateTypeUpper As String = UpdateType.ToUpper

                Dim IndexRow As Integer = Convert.ToInt32(e.CommandArgument)

                Dim Reason As String = ddlReason.SelectedValue

                If Reason = "" Then
                    lblError.Text = "Alasan tidak boleh kosong!"
                    GoTo Skip
                End If

                Dim dt As DataTable = ViewState("gvData")
                If Not IsNothing(dt) Then
                    If dt.Rows.Count > 0 Then

                        Dim TrackNum As String = dt.Rows(IndexRow).Item("TrackNum")

                        Dim dtUpdate As DataTable = bypassdelimanpin_update(TrackNum, UpdateTypeUpper, Reason, ErrorMessage)

                        If Not IsNothing(dtUpdate) Then
                            If dtUpdate.Rows.Count > 0 Then

                                Dim ResulValue As String = "" & dtUpdate.Rows(0).Item("ResultValue")

                                If ResulValue.StartsWith("ERROR-") Then

                                    Dim Keterangan As String = ResulValue.Substring(6, ResulValue.Length - 6)

                                    lblError.Text = "Update PIN " & UpdateType & " gagal! - " & Keterangan

                                Else

                                    Dim UpdateTypeName As String = "PIN" & UpdateType

                                    dt.Rows(IndexRow).Item(UpdateTypeName) = ResulValue
                                    dt.Rows(IndexRow).Item("BasePIN") = "" & dtUpdate.Rows(0).Item("BasePIN")

                                    ViewState("gvData") = dt

                                    gvData.DataSource = dt
                                    gvData.DataBind()

                                    lblError.Text = "Berhasil!"

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

    Private Function bypassdelimanpin_get(ByVal NoAWB As String, ByRef ErrorMessage As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim SqlQuery As String = ""
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = "call `bypassdelimanpin_get`("
            SqlQuery &= "@NoAWB"
            SqlQuery &= ")"

            Dim SqlParam As New Dictionary(Of String, String)
            SqlParam.Add("@NoAWB", NoAWB)

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

    Private Function bypassdelimanpin_update(ByVal TrackNum As String, ByVal UpdateType As String, ByVal Reason As String, ByRef ErrorMessage As String) As DataTable

        Dim dt As DataTable = Nothing

        If Not IsNothing(Session("NIK")) Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing

            Try

                SqlQuery = "call `bypassdelimanpin_update`(@TrackNum,@UpdateType,@Reason,@NIK);"

                Dim SqlParam As New Dictionary(Of String, String)
                SqlParam.Add("@TrackNum", TrackNum)
                SqlParam.Add("@UpdateType", UpdateType)
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
