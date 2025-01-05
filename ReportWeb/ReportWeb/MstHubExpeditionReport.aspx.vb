
Imports System.Data
Imports System.IO
Imports System.Collections.Generic
Imports ReportWeb.ClsWebVer
Imports MySql.Data.MySqlClient

Partial Class MstHubExpeditionReport
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

            dt.TableName = ""
            dt = ObjFungsi.GetHubList(Me.GetType().Name, Session("UserCode"))
            If dt.TableName.ToUpper <> "ERROR" Then
                ObjFungsi.ddlBound(ddlHub, dt, "Display", "Value")

                dt.DefaultView.RowFilter = "Value <> ''"
                dt = dt.DefaultView.ToTable
                ObjFungsi.chkBound(chkList1, dt, "Display", "Value")
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

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList2, dt, "Display", "Value")

            dt.TableName = ""
            dt = GetExpeditionList()
            ObjFungsi.ddlBound(ddlEkspedisi, dt, "Display", "Value")
            If dt.TableName.ToUpper <> "ERROR" Then
                ObjFungsi.ddlBound(ddlEkspedisi, dt, "Display", "Value")

                dt.DefaultView.RowFilter = "Value <> ''"
                dt = dt.DefaultView.ToTable
                ObjFungsi.chkBound(chkList3, dt, "Display", "Value")
            Else
                ErrorList.AppendLine("ERROR GetExpeditionList!")
            End If

            'ViewState("ddlHub") = ""
            'ViewState("ddlHubName") = ""

            'ViewState("ddlProses") = ""
            'ViewState("ddlProsesName") = ""

            'ViewState("ddlEkspedisi") = ""
            'ViewState("ddlEkspedisiName") = ""

            lblError.Text = ErrorList.ToString.Replace(vbCrLf, "<br/>")

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Protected Sub btnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProses.Click

        ProsesPreview(False)

    End Sub

    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreview.Click

        ProsesPreview(True)

    End Sub

    Private Function ValidasiProses() As Boolean

        Try

            lblError.Text = ""

            'ViewState("ddlHub") = ddlAsal.SelectedValue
            'ViewState("ddlHubName") = ddlAsal.SelectedItem.Text

            'ViewState("ddlProses") = ddlProses.SelectedValue
            'ViewState("ddlProsesName") = ddlProses.SelectedItem.Text

            'ViewState("ddlEkspedisi") = ddlEkspedisi.SelectedValue
            'ViewState("ddlEkspedisiName") = ddlEkspedisi.SelectedItem.Text

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

    Private Sub ProsesPreview(ByVal PreviewOnly As Boolean)

        Try

            If ValidasiProses() Then

                Dim Ekspedisi As String = GetEkspedisi("CODE")
                Dim EkspedisiName As String = GetEkspedisi("NAME")
                Dim HubAsal As String = GetHub("CODE")
                Dim HubAsalName As String = GetHub("NAME")
                Dim ProsesValue As String = GetProcess("CODE")
                Dim ProsesName As String = GetProcess("NAME")

                Proses(Ekspedisi, EkspedisiName, HubAsal, HubAsalName, ProsesValue, ProsesName, PreviewOnly)

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Sub Proses(ByVal Ekspedisi As String, ByVal NamaEkspedisi As String, ByVal HubAsal As String, ByVal HubAsalName As String, ByVal ProsesValue As String, ByVal ProsesName As String, ByVal PreviewOnly As Boolean)

        Dim User As String = "" & Session("NIK")

        If User <> "" Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing
            Dim ds As DataSet = Nothing

            Try

                SqlQuery = "call `report_msthubexpedition_compile`(@Hub,@ProsesValue,@Ekspedisi,@PreviewOnly);"

                Dim SqlParam As New Dictionary(Of String, String)
                SqlParam.Add("@Hub", HubAsal)
                SqlParam.Add("@ProsesValue", ProsesValue)
                SqlParam.Add("@Ekspedisi", Ekspedisi)
                SqlParam.Add("@PreviewOnly", PreviewOnly)

                MCon = ObjCon.SetConn_Slave1
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If
                ds = ObjSQL.SQLInsertIntoDataset(MCon, SqlQuery, SqlParam)

                Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

                If PreviewOnly Then

                    Dim HeaderRow As Integer = 2
                    Dim HeaderTitle(HeaderRow) As String
                    Dim HeaderContent(HeaderRow) As String
                    HeaderTitle(0) = "HUB ASAL"
                    HeaderTitle(1) = "PROSES"
                    HeaderTitle(2) = "EKSPEDISI"


                    HeaderContent(0) = HubAsalName
                    HeaderContent(1) = ProsesValue
                    HeaderContent(2) = NamaEkspedisi

                    Session("TitleReport") = "MstHub Expedition Report"
                    Session("HeaderTitleReport") = HeaderTitle
                    Session("HeaderContentReport") = HeaderContent
                    Session("BodyReportDs") = ds

                    'Response.Write("<script language=javascript>child=window.open('Preview.aspx');</script>")
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Preview_" & AutoNum, "child=window.open('Preview.aspx');", True)

                Else

                    Dim FileName As String = "SettingMstHubExpedition_" & AutoNum
                    Dim FileExt As String = ".csv"
                    Dim FileNameExt As String = FileName & FileExt

                    Dim PesanError As String = ""

                    Dim hasil As String = ObjFungsi.DataSetToFile(ds, "|", False, FileName, FileExt, User, PesanError)

                    Dim Result() As String = hasil.Split("|")
                    If Result(0) = "1" Then

                        Dim FileToDownload As String = "" & Result(1)

                        If FileToDownload <> "" Then

                            Dim fileInfo As FileInfo = New FileInfo(FileToDownload)

                            If fileInfo.Exists Then

                                Session("DownloadPage_FileToDownload") = FileToDownload
                                Session("DownloadPage_DLFileName") = FileName
                                Session("DownloadPage_DLFileExt") = FileExt

                                ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Download_" & AutoNum, "child=window.open('DownloadPage.aspx');", True)

                            Else

                                ObjFungsi.WriteTracelogTxt("File " & FileToDownload & " not exists")
                                lblError.Text = "File " & FileNameExt & " not exists"

                            End If

                        Else

                            ObjFungsi.WriteTracelogTxt("File " & FileNameExt & " fail to create!")
                            lblError.Text = "File " & FileNameExt & " fail to create!"

                        End If

                    Else
                        lblError.Text = Result(1)
                    End If

                End If

            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
                ObjFungsi.WriteTracelogTxt(Pesan)

                lblError.Text = ex.Message

            Finally

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

            End Try

        Else

            Response.Redirect("Login.aspx", False)

            Dim VirtualPath As String = Request.CurrentExecutionFilePath
            Dim FileName As String = System.Web.VirtualPathUtility.GetFileName(VirtualPath)
            Session("RedirectMessage") = FileName & ", User Kosong"

        End If

    End Sub

    '#Region "Button Multi Select"
    Private Function GetHub(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList1)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
                ScriptManager1.SetFocus(lblError)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlHub.SelectedItem.Text
                    Else
                        hasil1 = ddlHub.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_HubName")
                    Else
                        hasil1 = ViewState("Multi_Hub")
                    End If
                End If
            End If

            Return hasil1

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message
            ScriptManager1.SetFocus(lblError)

            Return Nothing

        End Try

        Return hasil1

    End Function

    Private Function GetProcess(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList2)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
                ScriptManager1.SetFocus(lblError)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlProses.SelectedItem.Text
                    Else
                        hasil1 = ddlProses.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_ProsesName")
                    Else
                        hasil1 = ViewState("Multi_Proses")
                    End If
                End If
            End If

            Return hasil1

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message
            ScriptManager1.SetFocus(lblError)

            Return Nothing

        End Try

        Return hasil1

    End Function

    Private Function GetEkspedisi(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList3)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlEkspedisi.SelectedItem.Text
                    Else
                        hasil1 = ddlEkspedisi.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_EkspedisiName")
                    Else
                        hasil1 = ViewState("Multi_Ekspedisi")
                    End If
                End If
            End If

            Return hasil1

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

            Return Nothing

        End Try

        Return hasil1

    End Function

    Protected Sub btnMulti1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMulti1.Click, btnMulti2.Click, btnMulti3.Click

        lblError.Text = ""

        Dim btn As Button = sender
        'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
        Dim btnID As String = (btn.ID).Replace("btnMulti", "")

        If btnID = "1" Then
            td1.Visible = False
            tr1.Visible = True
            ObjFungsi.SetMulti(chkList1, ViewState("Multi_Hub"))
        ElseIf btnID = "2" Then
            td2.Visible = False
            tr2.Visible = True
            ObjFungsi.SetMulti(chkList2, ViewState("Multi_Proses"))
        ElseIf btnID = "3" Then
            td3.Visible = False
            tr3.Visible = True
            ObjFungsi.SetMulti(chkList3, ViewState("Multi_Ekspedisi"))
        Else
        End If

    End Sub

    Protected Sub btnAll1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAll1.Click, btnAll2.Click, btnAll3.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnAll", "")

            If btnID = "1" Then
                ObjFungsi.SelectAll(chkList1)
            ElseIf btnID = "2" Then
                ObjFungsi.SelectAll(chkList2)
            ElseIf btnID = "3" Then
                ObjFungsi.SelectAll(chkList3)
            Else
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message
            ScriptManager1.SetFocus(lblError)

        End Try

    End Sub

    Protected Sub btnNon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNon1.Click, btnNon2.Click, btnNon3.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnNon", "")

            If btnID = "1" Then
                ObjFungsi.DeSelectAll(chkList1)
            ElseIf btnID = "2" Then
                ObjFungsi.DeSelectAll(chkList2)
            ElseIf btnID = "3" Then
                ObjFungsi.DeSelectAll(chkList3)
            Else
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message
            ScriptManager1.SetFocus(lblError)

        End Try

    End Sub

    Protected Sub btnBatal1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBatal1.Click, btnBatal2.Click, btnBatal3.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnBatal", "")

            If btnID = "1" Then
                td1.Visible = True
                tr1.Visible = False
                ScriptManager1.SetFocus(btnMulti1)
            ElseIf btnID = "2" Then
                td2.Visible = True
                tr2.Visible = False
                ScriptManager1.SetFocus(btnMulti2)
            ElseIf btnID = "3" Then
                td3.Visible = True
                tr3.Visible = False
                ScriptManager1.SetFocus(btnMulti3)
            Else
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message
            ScriptManager1.SetFocus(lblError)

        End Try

    End Sub

    Protected Sub btnPilih1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPilih1.Click, btnPilih2.Click, btnPilih3.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnPilih", "")

            If btnID = "1" Then

                td1.Visible = True
                tr1.Visible = False

                ViewState("Multi_Hub") = ObjFungsi.SetDataMulti(btnMulti1, chkList1)
                ViewState("Multi_HubName") = ObjFungsi.SetDataMultiName(btnMulti1, chkList1)

                ScriptManager1.SetFocus(btnMulti1)

            ElseIf btnID = "2" Then

                td2.Visible = True
                tr2.Visible = False

                ViewState("Multi_Proses") = ObjFungsi.SetDataMulti(btnMulti2, chkList2)
                ViewState("Multi_ProsesName") = ObjFungsi.SetDataMultiName(btnMulti2, chkList2)

                ScriptManager1.SetFocus(btnMulti2)

            ElseIf btnID = "3" Then

                td3.Visible = True
                tr3.Visible = False

                ViewState("Multi_Ekspedisi") = ObjFungsi.SetDataMulti(btnMulti3, chkList3)
                ViewState("Multi_EkspedisiName") = ObjFungsi.SetDataMultiName(btnMulti3, chkList3)

                ScriptManager1.SetFocus(btnMulti3)

            Else
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message
            ScriptManager1.SetFocus(lblError)

        End Try

    End Sub

    '#End Region

#Region "webservice"
    Private Function GetExpeditionList() As DataTable

        Dim dt As DataTable = Nothing

        dt = ObjFungsi.GetExpeditionList(Me.GetType().Name, Session("UserCode"))

        If dt.TableName = "ERROR" Then
            lblError.Text = dt.Rows(0).Item("RESPON")
        End If

        Return dt

    End Function
#End Region

End Class
