
Imports System.Data
Imports System.Collections.Generic
Imports MySql.Data.MySqlClient
Imports System.IO
Imports Microsoft.VisualBasic.ApplicationServices
Imports ReportWeb.CoreService
Imports ReportWeb.ClsWebVer

Partial Class CalculateHubMustReport
    Inherits System.Web.UI.Page

    Dim MyPage As String
    Dim ObjFungsi As New ClsFungsi
    Dim ObjCon As New ClsConnection
    Dim ObjSQL As New ClsSQL
    Dim serv As New LocalCore
    Dim param() As Object
    Dim respon() As Object
    Dim ScriptManager1 As New ScriptManager


    Dim SqlParam As New Dictionary(Of String, String)

    Dim PageTimeout As Integer = 3600000 'In miliseconds

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("NIK") = "2015548000"
        Session("NameOfUser") = "Ryu"

        ScriptManager1 = ScriptManager.GetCurrent(Me.Page)
        'PageTimeout in miliseconds
        ScriptManager1.AsyncPostBackTimeout = PageTimeout / 1000 'In Seconds

        MyPage = System.Web.VirtualPathUtility.GetFileName(Request.RawUrl).Replace(".aspx", "")
        Dim MyLi As String = "li" & MyPage

        Dim AllowedMenu As String = "" & Session("AllowedMenu")
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

        Dim MCon As MySqlConnection = Nothing
        Dim dt As New DataTable
        Dim ErrorList As New StringBuilder
        Dim PesanError As String = ""
        Dim dsData As New DataSet

        Try
            MCon = ObjCon.SetConn_Slave1

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjFungsi.GetECommerceList(Me.GetType().Name, Session("UserCode"))
            If dt.TableName.ToUpper <> "ERROR" Then

                ObjFungsi.ddlBound(ddlECom, dt, "Display", "Value")
                dt.DefaultView.RowFilter = "Value <> ''"

                dt = dt.DefaultView.ToTable
                ObjFungsi.chkBound(chkList51, dt, "Display", "Value")

            Else
                ErrorList.AppendLine("ERROR GetECommerceList!")
            End If

            Dim ErrorMessage As String = ""

            dt = GetDCList()
            If dt.TableName.ToUpper <> "ERROR" Then

                ObjFungsi.ddlBound(ddlCabangDc, dt, "Display", "Value")
                dt.DefaultView.RowFilter = "Value <> ''"

                dt = dt.DefaultView.ToTable
                ObjFungsi.chkBound(chkList2, dt, "Display", "Value")

            Else
                ErrorList.AppendLine("ERROR GetDcList!")
            End If

            ErrorMessage = ""
            dt = GetHubList()
            If dt.TableName.ToUpper <> "ERROR" Then

                ObjFungsi.ddlBound(ddlHub, dt, "Display", "Value")
                dt.DefaultView.RowFilter = "Value <> ''"

                dt = dt.DefaultView.ToTable
                ObjFungsi.chkBound(chkList4, dt, "Display", "Value")

            Else
                ErrorList.AppendLine("ERROR GetHubList!")
            End If


            lblError.Text = ErrorList.ToString.Replace(vbCrLf, "<br/>")

            ViewState("TanggalAwal") = ""
            ViewState("TanggalAkhir") = ""
        Catch ex As Exception
            lblError.Text = ex.Message
            ScriptManager1.SetFocus(TxtError)

        Finally
            If Not MCon Is Nothing Then
                If MCon.State <> ConnectionState.Closed Then
                    MCon.Close()
                End If
                MCon.Dispose()
            End If
        End Try

    End Sub

    Private Function ValidasiProses(ByRef NoToko As String, ByRef NoAwb As String) As Boolean

        Try

            lblError.Text = ""

            txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
            txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

            Dim ask_date As Boolean = True

            NoToko = txtKodeToko.Text.Trim.Replace(vbCrLf, "|")
            NoAwb = txtNoAwb.Text.Trim.Replace(vbCrLf, "|")
            Dim Toko() As String = NoToko.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)
            Dim Awb() As String = NoAwb.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)

            If Toko.Length > 20 Then
                lblError.Text = "Maksimal Jumlah Toko 20 !!"
                Return False
            End If

            If Awb.Length > 20 Then
                lblError.Text = "Maksimal Jumlah Nomor Resi 20 !!"
                Return False
            End If

            If Toko.Length <= 1 Then

                If Toko.Length = 1 Then
                    If Toko(0) <> "" Then
                        ask_date = False
                    End If
                End If

                If ask_date Then

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

                    Dim PesanError As String = ""
                    If ObjFungsi.LimitPeriode(txtTgl1.Text, txtTgl2.Text, PesanError) = False Then
                        lblError.Text = PesanError
                        Return False
                    End If

                End If

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

    Protected Sub btnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProses.Click

        ProsesPreview(False)

    End Sub

    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreview.Click

        ProsesPreview(True)

    End Sub

    Private Sub ProsesPreview(ByVal PreviewOnly As Boolean)

        Dim NoToko As String = ""
        Dim NoAwb As String = ""

        Try
            If ValidasiProses(NoToko, NoAwb) Then
                Dim TanggalAwal As String = txtTgl1.Text
                Dim TanggalAkhir As String = txtTgl2.Text
                Dim KodeEcommerce As String = GetEcommerce("CODE")
                Dim NamaEcommerce As String = GetEcommerce("NAME")
                Dim KodeHub As String = GetHub("CODE")
                Dim NamaHub As String = GetHub("NAME")
                Dim KodeDC As String = GetDc("CODE")
                Dim NamaDC As String = GetDc("NAME")

                Proses(TanggalAwal, TanggalAkhir, KodeEcommerce, NoToko, KodeDC, KodeHub, NoAwb, NamaEcommerce, NamaHub, NamaDC, PreviewOnly)

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Sub Proses(ByVal TanggalAwal As String, ByVal TanggalAkhir As String, ByVal KodeEcommerce As String, ByVal NoToko As String, ByVal KodeDC As String, ByVal KodeHub As String, ByVal NoAwb As String _
                       , ByVal NamaEcommerce As String, ByVal NamaHub As String, ByVal NamaDC As String, ByVal PreviewOnly As Boolean)

        Dim User As String = "" & Session("NIK")

        If User <> "" Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing
            Dim dt As DataTable = Nothing

            Try

                SqlQuery &= "call calculate_hubmustwgtdim("
                SqlQuery &= " @StartDate"
                SqlQuery &= " , @EndDate"
                SqlQuery &= " , @KodeEcommerce"
                SqlQuery &= " , @NoToko"
                SqlQuery &= " , @DcList"
                SqlQuery &= " , @HubList"
                SqlQuery &= " , @TrackNum"
                SqlQuery &= " , @Preview"
                SqlQuery &= " )"

                Dim SqlParam As New Dictionary(Of String, String)
                SqlParam.Add("@StartDate", TanggalAwal)
                SqlParam.Add("@EndDate", TanggalAkhir)
                SqlParam.Add("@KodeEcommerce", KodeEcommerce)
                SqlParam.Add("@NoToko", NoToko)
                SqlParam.Add("@HubList", KodeHub)
                SqlParam.Add("@DcList", KodeDC)
                SqlParam.Add("@TrackNum", NoAwb)
                SqlParam.Add("@Preview", PreviewOnly)

                MCon = ObjCon.SetConn_Slave1
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

                Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

                If PreviewOnly Then

                    Dim HeaderRow As Integer = 6
                    Dim HeaderTitle(HeaderRow) As String
                    Dim HeaderContent(HeaderRow) As String

                    HeaderTitle(0) = "TANGGAL AWAL"
                    HeaderTitle(1) = "TANGGAL AKHIR"
                    HeaderTitle(2) = "Partner"
                    HeaderTitle(3) = "Hub Asal"
                    HeaderTitle(4) = "DC Asal"
                    HeaderTitle(5) = "Toko Asal"
                    HeaderTitle(6) = "AWB"

                    HeaderContent(0) = TanggalAwal
                    HeaderContent(1) = TanggalAkhir
                    HeaderContent(2) = NamaEcommerce
                    HeaderContent(3) = NamaHub
                    HeaderContent(4) = NamaDC
                    HeaderContent(5) = NoToko
                    HeaderContent(6) = NoAwb

                    Session("TitleReport") = "LAPORAN TIMBANG UKUR ULANG"
                    Session("HeaderTitleReport") = HeaderTitle
                    Session("HeaderContentReport") = HeaderContent
                    Session("BodyReport") = dt

                    'Response.Write("<script language=javascript>child=window.open('Preview.aspx');</script>")
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Preview_" & AutoNum, "child=window.open('Preview.aspx');", True)

                Else

                    Dim FileName As String = "TimbangUkurUlangReport_" & AutoNum
                    Dim FileExt As String = ".csv"
                    Dim FileNameExt As String = FileName & FileExt

                    Dim PesanError As String = ""

                    Dim hasil As String = ObjFungsi.DataTableToFile(dt, "|", False, FileName, FileExt, User, PesanError)

                    Dim Result() As String = hasil.Split("|")
                    If Result(0) = "1" Then

                        Dim FileToDownload As String = Result(1)

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

#Region "Button Multi Select"

    Private Function GetEcommerce(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList51)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlECom.SelectedItem.Text
                    Else
                        hasil1 = ddlECom.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_EcommName")
                    Else
                        hasil1 = ViewState("Multi_Ecomm")
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

    Private Function GetDCList() As DataTable

        Dim dt As New DataTable
        dt = Nothing

        ReDim param(3)
        param(0) = UserWS
        param(1) = PassWS
        param(2) = "" 'DcCode
        param(3) = "IDM"

        respon = serv.GetDCList(AppName, AppVersion, param)

        If respon(0) = "0" Then
            dt = ObjFungsi.ConvertStringToDatatable(respon(2).ToString)
        Else
            lblError.Text = respon(1).ToString
        End If

        Return dt

    End Function

    Private Function GetHubList() As DataTable

        Dim dt As DataTable = Nothing

        dt = ObjFungsi.GetHubList(Me.GetType().Name, Session("UserCode"))

        If dt.TableName = "ERROR" Then
            lblError.Text = dt.Rows(0).Item("RESPON")
        End If

        Return dt

    End Function

    Private Function GetHub(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList4)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
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

            Return Nothing

        End Try

        Return hasil1

    End Function

    Private Function GetDc(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList2)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlCabangDc.SelectedItem.Text
                    Else
                        hasil1 = ddlCabangDc.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_CabangDCName")
                    Else
                        hasil1 = ViewState("Multi_CabangDC")
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


    Protected Sub btnMulti1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMulti2.Click, btnMulti4.Click, btnMulti51.Click

        lblError.Text = ""

        txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
        txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

        Dim btn As Button = sender
        'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
        Dim btnID As String = (btn.ID).Replace("btnMulti", "")

        If btnID = "2" Then
            td2.Visible = False
            tr2.Visible = True
            ObjFungsi.SetMulti(chkList2, ViewState("Multi_CabangDC"))
        ElseIf btnID = "4" Then
            td4.Visible = False
            tr4.Visible = True
            ObjFungsi.SetMulti(chkList4, ViewState("Multi_Hub"))
        ElseIf btnID = "51" Then
            td51.Visible = False
            tr51.Visible = True
            ObjFungsi.SetMulti(chkList51, ViewState("Multi_Ecomm"))
        Else
        End If

    End Sub

    Protected Sub btnAll1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAll2.Click, btnAll4.Click, btnAll51.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnAll", "")

            If btnID = "2" Then
                ObjFungsi.SelectAll(chkList2)
            ElseIf btnID = "4" Then
                ObjFungsi.SelectAll(chkList4)
            ElseIf btnID = "51" Then
                ObjFungsi.SelectAll(chkList51)

            Else
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Protected Sub btnNon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNon2.Click, btnNon4.Click, btnNon51.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnNon", "")

            If btnID = "2" Then
                ObjFungsi.DeSelectAll(chkList2)
            ElseIf btnID = "4" Then
                ObjFungsi.DeSelectAll(chkList4)
            ElseIf btnID = "51" Then
                ObjFungsi.DeSelectAll(chkList51)
            Else
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Protected Sub btnBatal1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBatal2.Click, btnBatal4.Click, btnBatal51.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnBatal", "")

            If btnID = "2" Then
                td2.Visible = True
                tr2.Visible = False
                ScriptManager1.SetFocus(btnMulti2)
            ElseIf btnID = "4" Then
                td4.Visible = True
                tr4.Visible = False
                ScriptManager1.SetFocus(btnMulti4)
            ElseIf btnID = "51" Then
                td51.Visible = True
                tr51.Visible = False
                ScriptManager1.SetFocus(btnMulti51)
            Else
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Protected Sub btnPilih1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPilih2.Click, btnPilih4.Click, btnPilih51.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnPilih", "")

            If btnID = "2" Then

                td2.Visible = True
                tr2.Visible = False

                ViewState("Multi_CabangDC") = ObjFungsi.SetDataMulti(btnMulti2, chkList2)
                ViewState("Multi_CabangDCName") = ObjFungsi.SetDataMultiName(btnMulti2, chkList2)

                ScriptManager1.SetFocus(btnMulti2)

            ElseIf btnID = "4" Then

                td4.Visible = True
                tr4.Visible = False

                ViewState("Multi_Hub") = ObjFungsi.SetDataMulti(btnMulti4, chkList4)
                ViewState("Multi_HubName") = ObjFungsi.SetDataMultiName(btnMulti4, chkList4)

                ScriptManager1.SetFocus(btnMulti4)

            ElseIf btnID = "51" Then

                td51.Visible = True
                tr51.Visible = False

                ViewState("Multi_Ecomm") = ObjFungsi.SetDataMulti(btnMulti51, chkList51)
                ViewState("Multi_EcommName") = ObjFungsi.SetDataMultiName(btnMulti51, chkList51)

                ScriptManager1.SetFocus(btnMulti51)
            Else
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

#End Region

End Class
