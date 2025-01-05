
Imports System.Data
Imports System.IO
Imports ReportWeb.ClsWebVer
Imports MySql.Data.MySqlClient

Partial Class SettingMstHubExpedition
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

        'Dim MyLink As String = "Link" & MyPage
        'Session("CurrentMenu") = MyLink

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

            dt.TableName = ""
            dt = ObjFungsi.GetHubList(Me.GetType().Name, Session("UserCode"))
            If dt.TableName.ToUpper <> "ERROR" Then
                ObjFungsi.ddlBound(ddlAsal, dt, "Display", "Value")
            Else
                ErrorList.AppendLine("ERROR GetHubList!")
            End If

            dt = New DataTable
            dt.Columns.Add("Value")
            dt.Columns.Add("Display")

            dt.Rows.Add(New String() {"", ""})
            dt.Rows.Add(New String() {"SERAH", "SERAH"})
            dt.Rows.Add(New String() {"TERIMA", "TERIMA"})
            dt.Rows.Add(New String() {"JEMPUT REKANAN", "JEMPUT REKANAN"})
            dt.Rows.Add(New String() {"JEMPUT EKSPEDISI", "JEMPUT EKSPEDISI"})
            dt.Rows.Add(New String() {"SEWA ARMADA", "SEWA ARMADA"})
            ObjFungsi.ddlBound(ddlProses, dt, "Display", "Value")

            lblError.Text = ErrorList.ToString.Replace(vbCrLf, "<br/>")

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

            ViewState("old_ddlAsal") = ""
            ViewState("old_ddlAsalName") = ""

            ViewState("old_ddlProses") = ""
            ViewState("old_ddlProsesName") = ""

            ViewState("gvData") = Nothing
            gvData.DataSource = Nothing
            gvData.DataBind()

            TrData.Visible = False
            TrSearch.Visible = False
            lblSearch.Text = ""

            If ddlAsal.SelectedValue = "" Then
                lblError.Text = lblAsal.Text & " belum dipilih !"
                Return False
            End If

            If ddlProses.SelectedValue = "" Then
                lblError.Text = lblProses.Text & " belum dipilih !"
                Return False
            End If

            ViewState("old_ddlAsal") = ddlAsal.SelectedValue
            ViewState("old_ddlAsalName") = ddlAsal.SelectedItem.Text

            ViewState("old_ddlProses") = ddlProses.SelectedValue
            ViewState("old_ddlProsesName") = ddlProses.SelectedItem.Text

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

        If Not IsNothing(Session("NIK")) Then

            If validasiCari() Then

                Dim dsData As DataSet = Nothing

                Dim PesanError As String = ""

                Dim Hub As String = ViewState("old_ddlAsal")
                Dim Proses As String = ViewState("old_ddlProses")

                Dim dt As DataTable = ObjFungsi.GetHubProcessTypeExpedition(Hub, Proses, dsData, PesanError)

                If Not IsNothing(dt) Then

                    If dt.Rows.Count > 0 Then

                        gvData.DataSource = dt
                        gvData.DataBind()

                        ViewState("gvData") = dt

                        TrData.Visible = True
                        TrSearch.Visible = True

                        Dim HubName As String = ViewState("old_ddlAsalName")
                        Dim ProsesName As String = ViewState("old_ddlProsesName")

                        dt.DefaultView.RowFilter = "Status = '1'"

                        'Hub = HXXX, Proses = YYYY, Ekspedisi Aktif = NRecord
                        lblSearch.Text = "Hub = " & HubName & ", Proses = " & ProsesName & ", Ekspedisi Aktif = " & dt.DefaultView.Count

                        dt.DefaultView.RowFilter = Nothing

                    Else
                        lblError.Text = lblDaftarEkspedisi.Text & " Kosong!"
                    End If

                Else
                    lblError.Text = PesanError
                End If

            End If

        End If

    End Sub

    Private Function GetDataGrid(ByRef ErrorMessage As String) As String

        Dim Hasil As String = ""
        'Account1|Account2|... (hanya yang aktif)

        Try

            Dim dt As DataTable = ViewState("gvData")

            If Not IsNothing(dt) Then
                If dt.Rows.Count > 0 Then

                    Dim SB As New StringBuilder

                    Dim CBStatus As CheckBox

                    Dim index As Integer = 0

                    For i As Integer = 0 To gvData.Rows.Count - 1

                        CBStatus = DirectCast(Me.gvData.Rows(i).FindControl("CBStatus"), CheckBox)

                        If CBStatus.Checked Then

                            If index > 0 Then
                                SB.Append("|")
                            End If

                            SB.Append(dt.Rows(i).Item("Account"))

                            index += 1

                        End If

                    Next

                    Hasil = SB.ToString

                Else
                    ErrorMessage = "Tidak ada data di Grid!"
                End If
            Else
                ErrorMessage = "Tidak ada data di Grid!"
            End If

        Catch ex As Exception

            ErrorMessage = ex.Message

        End Try

        Return Hasil

    End Function

    Private Function validasiSimpan(ByRef ExpeditionList As String) As Boolean

        Try

            lblError.Text = ""

            If ddlAsal.SelectedValue <> ViewState("old_ddlAsal") Then
                lblError.Text = lblAsal.Text & " berubah tekan tombol Cari kembali!"
                Return False
            End If

            If ddlProses.SelectedValue <> ViewState("old_ddlProses") Then
                lblError.Text = lblProses.Text & " berubah tekan tombol Cari kembali!"
                Return False
            End If

            Dim ErrorMessage As String = ""
            ExpeditionList = GetDataGrid(ErrorMessage)
            If ErrorMessage <> "" Then
                lblError.Text = ErrorMessage
                Return False
            End If

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

    Protected Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim UserId As String = "" & Session("NIK")

        If UserId <> "" Then

            Dim ExpeditionList As String = ""

            If validasiSimpan(ExpeditionList) Then

                Dim HubCode As String = ViewState("old_ddlAsal")
                Dim ProcessType As String = ViewState("old_ddlProses")

                If UserId <> "" Then

                    Dim ErrorMessage As String = ""

                    If ObjFungsi.SetHubProcessTypeExpedition(HubCode, ProcessType, ExpeditionList, UserId, ErrorMessage) Then
                        Session("ResultMessage" & MyPage) = "Berhasil!"
                        Response.Redirect(Request.RawUrl, False)
                    Else
                        lblError.Text = ErrorMessage
                    End If

                End If

            End If

        End If

    End Sub

End Class
