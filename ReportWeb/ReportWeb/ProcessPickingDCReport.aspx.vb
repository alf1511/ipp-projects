
Imports System.Data
Imports System.IO
Imports System.Collections.Generic
Imports ReportWeb.ClsWebVer
Imports MySql.Data.MySqlClient
Imports System.Drawing

Partial Class ProcessPickingDCReport
    Inherits System.Web.UI.Page

    Dim MyPage As String
    Dim ObjFungsi As New ClsFungsi
    Dim ObjCon As New ClsConnection
    Dim ObjSQL As New ClsSQL
    Dim ObjService As New ClsService
    Dim serv As New LocalCore
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

        Dim ErrorList As New StringBuilder
        Dim dt As New DataTable

        dt.TableName = ""
        Dim ErrorMessage As String = ""

        ErrorMessage = ""
        dt = GetHubList()
        If dt.TableName.ToUpper <> "ERROR" Then

            ObjFungsi.ddlBound(ddlHub, dt, "Display", "Value")
            dt.DefaultView.RowFilter = "Value <> ''"

            'dt = dt.DefaultView.ToTable
            'ObjFungsi.chkBound(chkList4, dt, "Display", "Value")

        Else
            ErrorList.AppendLine("ERROR GetHubList!")
        End If

        lblError.Text = ErrorList.ToString.Replace(vbCrLf, "<br/>")

    End Sub

    Private Function validasiProses(ByRef Kode As String) As Boolean

        Try

            lblError.Text = ""

            txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
            txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

            'cek tanggal awal
            If txtTgl1.Text = "" Then
                lblError.Text = "Tanggal Awal belum dipilih !!"
                Return False
            End If

            'cek tanggal akhir
            If txtTgl2.Text = "" Then
                lblError.Text = "Tanggal Akhir belum dipilih !!"
                Return False
            End If

            Kode = TxtKodeToko.Text.Trim.Replace(vbCrLf, "/")
            'If TxtKodeToko.Text = "" Then
            '    lblError.Text = lblKodeToko.Text & " masih kosong!"
            '    Return False
            'End If
            Dim KodeToko() As String = Kode.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)
            ViewState("HubCode") = ddlHub.SelectedValue
            ViewState("HubName") = ddlHub.SelectedItem.Text


            Dim PesanError As String = ""
            If ObjFungsi.LimitPeriode(txtTgl1.Text, txtTgl2.Text, PesanError) = False Then
                lblError.Text = PesanError
                Return False
            End If

            'If ddlECom.SelectedValue = "" Then
            '    lblError.Text = LblECom.Text & " belum dipilih !!"
            '    Return False
            'End If

            'Dim Ecomm As String = GetEcomm("CODE")
            'Dim EcommList As String() = Ecomm.Split("|")
            'If EcommList.Length > 10 Then
            '    lblError.Text = "Jumlah maksimal Ecommerce adalah 10, silahkan periksa kembali!"
            '    ScriptManager1.SetFocus(lblError)
            '    Return False
            'End If

            Return True

        Catch ex As Exception

            lblError.Text = ex.Message
            Return False

        End Try

    End Function

    Private Sub ProsesPreview(ByVal PreviewOnly As Boolean)
        Dim KodeToko As String = ""
        If validasiProses(KodeToko) Then

            Dim TanggalAwal As String = txtTgl1.Text
            Dim TanggalAkhir As String = txtTgl2.Text
            Dim KodeHub As String = ViewState("HubCode")
            Dim NamaHub As String = ViewState("HubName")

            Proses(TanggalAwal, TanggalAkhir, KodeHub, NamaHub, KodeToko, PreviewOnly)

        End If

    End Sub

    Private Sub Proses(ByVal TanggalAwal As String, ByVal TanggalAkhir As String, ByVal KodeHub As String, ByVal NamaHub As String, ByVal KodeToko As String, ByVal PreviewOnly As Boolean)

        Dim User As String = "" & Session("NIK")

        If User <> "" Then

            'PageTimeout in miliseconds
            Server.ScriptTimeout = PageTimeout / 1000 'In Seconds
            Session.Timeout = PageTimeout / 1000 / 60 'In Minutes

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing
            Dim dt As DataTable = Nothing
            Dim ds As DataSet = Nothing
            Try

                SqlQuery = "call `report_dcrequest`(@TanggalAwal,@TanggalAkhir,@KodeHub,@KodeToko,@PreviewOnly);"

                Dim SqlParam As New Dictionary(Of String, String)
                SqlParam.Add("@TanggalAwal", TanggalAwal)
                SqlParam.Add("@TanggalAkhir", TanggalAkhir)
                SqlParam.Add("@KodeHub", KodeHub)
                SqlParam.Add("@NamaHub", NamaHub)
                SqlParam.Add("@KodeToko", KodeToko)
                SqlParam.Add("@PreviewOnly", PreviewOnly)

                MCon = ObjCon.SetConn_Slave1
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If
                dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

                If PreviewOnly Then

                    Dim HeaderRow As Integer = 3
                    Dim HeaderTitle(HeaderRow) As String
                    Dim HeaderContent(HeaderRow) As String
                    HeaderTitle(0) = "TANGGAL AWAL"
                    HeaderTitle(1) = "TANGGAL AKHIR"
                    HeaderTitle(2) = "NAMA HUB"
                    HeaderTitle(3) = "KODE TOKO"

                    HeaderContent(0) = TanggalAwal
                    HeaderContent(1) = TanggalAkhir
                    HeaderContent(2) = NamaHub
                    HeaderContent(3) = KodeToko

                    Session("TitleReport") = "LAPORAN DC PICKING"
                    Session("HeaderTitleReport") = HeaderTitle
                    Session("HeaderContentReport") = HeaderContent
                    Session("BodyReport") = dt

                    Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Preview_" & AutoNum, "child=window.open('Preview.aspx');", True)
                    'Jangan pakai Response.Write, font jadi besar2
                    'Response.Write("<script language=javascript>child=window.open('Preview.aspx');</script>")

                Else

                    Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

                    Dim FileName As String = "LapProsesDCPicking" & AutoNum
                    Dim FileExt As String = ".csv"
                    Dim FileNameExt As String = FileName & FileExt
                    Dim PesanError As String = ""

                    Dim hasil As String = ObjFungsi.DataTableToFile(dt, "|", False, FileName, FileExt, User, PesanError)

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
                                lblError.Text = "File " & FileToDownload & " not exists"

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

        End If

    End Sub

    Protected Sub btnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProses.Click

        ProsesPreview(False)

    End Sub

    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreview.Click

        ProsesPreview(True)

    End Sub

    Private Function GetHubList() As DataTable

        Dim dt As DataTable = Nothing

        dt = ObjFungsi.GetHubList(Me.GetType().Name, Session("UserCode"))

        If dt.TableName = "ERROR" Then
            lblError.Text = dt.Rows(0).Item("RESPON")
        End If

        Return dt

    End Function

    '#Region "Button Multi Select"

    '    Private Function GetEcomm(ByVal Tipe As String) As String

    '        Dim hasil1 As String = ""

    '        Tipe = Tipe.ToUpper

    '        Try

    '            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList51)

    '            If hasil2(0) = "9" Then
    '                lblError.Text = hasil2(1)
    '            Else
    '                If hasil2(0) = "0" Or hasil2(0) = "2" Then
    '                    If Tipe = "NAME" Then
    '                        hasil1 = ddlECom.SelectedItem.Text
    '                    Else
    '                        hasil1 = ddlECom.SelectedValue
    '                    End If
    '                Else
    '                    If Tipe = "NAME" Then
    '                        hasil1 = ViewState("Multi_EcommName")
    '                    Else
    '                        hasil1 = ViewState("Multi_Ecomm")
    '                    End If
    '                End If
    '            End If

    '            Return hasil1

    '        Catch ex As Exception

    '            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '            Dim Pesan As String = ""
    '            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
    '            ObjFungsi.WriteTracelogTxt(Pesan)

    '            lblError.Text = ex.Message

    '            Return Nothing

    '        End Try

    '        Return hasil1

    '    End Function

    '    Protected Sub btnMulti1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMulti51.Click

    '        lblError.Text = ""

    '        Try
    '            txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
    '            txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

    '            Dim btn As Button = sender
    '            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
    '            Dim btnID As String = (btn.ID).Replace("btnMulti", "")

    '            If btnID = "51" Then
    '                td51.Visible = False
    '                tr51.Visible = True
    '                ObjFungsi.SetMulti(chkList51, ViewState("Multi_Ecomm"))

    '            Else

    '            End If

    '        Catch ex As Exception
    '            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '            Dim Pesan As String = ""
    '            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
    '            ObjFungsi.WriteTracelogTxt(Pesan)

    '            lblError.Text = ex.Message

    '        End Try

    '    End Sub

    '    Protected Sub btnAll1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAll51.Click

    '        lblError.Text = ""

    '        Try
    '            Dim btn As Button = sender
    '            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
    '            Dim btnID As String = (btn.ID).Replace("btnAll", "")

    '            If btnID = "51" Then
    '                ObjFungsi.SelectAll(chkList51)
    '            Else

    '            End If

    '        Catch ex As Exception

    '            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '            Dim Pesan As String = ""
    '            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
    '            ObjFungsi.WriteTracelogTxt(Pesan)

    '            lblError.Text = ex.Message

    '        End Try

    '    End Sub

    '    Protected Sub btnNon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNon51.Click

    '        lblError.Text = ""

    '        Try
    '            Dim btn As Button = sender
    '            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
    '            Dim btnID As String = (btn.ID).Replace("btnNon", "")

    '            If btnID = "51" Then
    '                ObjFungsi.DeSelectAll(chkList51)
    '            Else

    '            End If

    '        Catch ex As Exception

    '            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '            Dim Pesan As String = ""
    '            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
    '            ObjFungsi.WriteTracelogTxt(Pesan)

    '            lblError.Text = ex.Message

    '        End Try

    '    End Sub

    '    Protected Sub btnBatal1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBatal51.Click

    '        lblError.Text = ""

    '        Try
    '            Dim btn As Button = sender
    '            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
    '            Dim btnID As String = (btn.ID).Replace("btnBatal", "")

    '            If btnID = "51" Then
    '                td51.Visible = True
    '                tr51.Visible = False
    '                ScriptManager1.SetFocus(btnMulti51)

    '            Else

    '            End If

    '        Catch ex As Exception

    '            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '            Dim Pesan As String = ""
    '            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
    '            ObjFungsi.WriteTracelogTxt(Pesan)

    '            lblError.Text = ex.Message

    '        End Try

    '    End Sub

    '    Protected Sub btnPilih1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPilih51.Click

    '        lblError.Text = ""

    '        Try
    '            Dim btn As Button = sender
    '            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
    '            Dim btnID As String = (btn.ID).Replace("btnPilih", "")

    '            If btnID = "51" Then
    '                td51.Visible = True
    '                tr51.Visible = False

    '                ViewState("Multi_Ecomm") = ObjFungsi.SetDataMulti(btnMulti51, chkList51)
    '                ViewState("Multi_EcommName") = ObjFungsi.SetDataMultiName(btnMulti51, chkList51)

    '                ScriptManager1.SetFocus(btnMulti51)

    '            Else

    '            End If

    '        Catch ex As Exception

    '            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '            Dim Pesan As String = ""
    '            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
    '            ObjFungsi.WriteTracelogTxt(Pesan)

    '            lblError.Text = ex.Message

    '        End Try

    '    End Sub

    ''#End Region

End Class
