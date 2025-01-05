Imports System.Data
Imports ReportWeb.ClsWebVer
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Collections.Generic
Imports Microsoft.VisualBasic.ApplicationServices
Imports System.Drawing
Public Class LaporanStatusMTORET
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

    Dim PageTimeout As Integer = 7200000 'In miliseconds
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
            Dim ErrorMessage As String = ""

            Dim dt As DataTable = GetDCList()
            ObjFungsi.ddlBound(ddlCabangDc, dt, "Display", "Value")

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList2, dt, "Display", "Value")


            dt = ObjFungsi.ConvertStringToDatatable("Display,Value|,|MTO,STI|RET,RTS")
            ObjFungsi.ddlBound(ddlStatus, dt, "Display", "Value")

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList1, dt, "Display", "Value")


            ViewState("TanggalAwal") = ""
            ViewState("TanggalAkhir") = ""

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Function GetDCList() As DataTable

        Dim dt As New DataTable
        dt = Nothing

        ReDim param(1)
        param(0) = UserWS
        param(1) = PassWS

        respon = serv.GetDCList(AppName, AppVersion, param)

        If respon(0) = "0" Then
            dt = ObjFungsi.ConvertStringToDatatable(respon(2).ToString)
        Else
            lblError.Text = respon(1).ToString
        End If

        Return dt

    End Function

    Private Function validasiProses(ByRef Kode As String) As Boolean

        Try

            lblError.Text = ""

            txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
            txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

            Kode = txtKodeToko.Text.Trim.Replace(vbCrLf, "/")
            Dim KodeToko() As String = Kode.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)

            If KodeToko.Length > 20 Then
                lblError.Text = "Maksimal 20 Toko !!"
                Return False
            End If

            If txtTgl1.Text = "" Then
                lblError.Text = "Tanggal Awal belum dipilih !!"
                Return False
            End If

            If txtTgl2.Text = "" Then
                lblError.Text = "Tanggal Akhir belum dipilih !!"
                Return False
            End If

            Dim PesanError As String = ""
            If ObjFungsi.LimitPeriode(txtTgl1.Text, txtTgl2.Text, PesanError) = False Then
                lblError.Text = PesanError
                Return False
            End If

            Return True

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Function

    Protected Sub btnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProses.Click

        ProsesPreview(False)

    End Sub

    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreview.Click

        ProsesPreview(True)

    End Sub

    Private Sub ProsesPreview(ByVal PreviewOnly As Boolean)

        Dim Kode As String = ""

        Try

            If validasiProses(Kode) Then

                Dim TanggalAwal As String = txtTgl1.Text
                Dim TanggalAkhir As String = txtTgl2.Text
                Dim KodeDc = GetDc("CODE")
                Dim NamaDC = GetDc("NAME")
                Dim Status As String = GetStatus("STATUS")

                Proses(KodeDc, Status, TanggalAwal, TanggalAkhir, Kode, NamaDC, PreviewOnly)

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Sub Proses(ByVal KodeDc As String, ByVal Status As String, ByVal TanggalAwal As String, ByVal TanggalAkhir As String, ByVal KodeToko As String, ByVal NamaDC As String, ByVal PreviewOnly As Boolean)

        Dim User As String = "" & Session("NIK")

        If User <> "" Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing
            Dim ds As DataSet = Nothing

            Try

                SqlQuery = "call `report_harian_toko_fokus_paket_todoor`(@TanggalAwal,@TanggalAkhir,@StatusCode,@KodeDc,@KodeToko,@PreviewOnly);"

                Dim SqlParam As New Dictionary(Of String, String)
                SqlParam.Add("@TanggalAwal", TanggalAwal)
                SqlParam.Add("@TanggalAkhir", TanggalAkhir)
                SqlParam.Add("@StatusCode", Status)
                SqlParam.Add("@KodeDc", KodeDc)
                SqlParam.Add("@KodeToko", KodeToko)
                SqlParam.Add("@PreviewOnly", PreviewOnly)

                MCon = ObjCon.SetConn_Slave1
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

                If PreviewOnly Then

                    Dim HeaderRow As Integer = 4
                    Dim HeaderTitle(HeaderRow) As String
                    Dim HeaderContent(HeaderRow) As String

                    HeaderTitle(0) = "TANGGAL AWAL"
                    HeaderTitle(1) = "TANGGAL AKHIR"
                    HeaderTitle(2) = "STATUS"
                    HeaderTitle(3) = "DC"
                    HeaderTitle(4) = "KODE TOKO"

                    HeaderContent(0) = TanggalAwal
                    HeaderContent(1) = TanggalAkhir
                    HeaderContent(2) = Status
                    HeaderContent(3) = NamaDC
                    HeaderContent(4) = KodeToko.Replace("|", " / ")

                    ds = ObjSQL.SQLInsertIntoDataset(MCon, SqlQuery, SqlParam)

                    Session("TitleReport") = "LAPORAN STATUS MTO/RET"
                    Session("HeaderTitleReport") = HeaderTitle
                    Session("HeaderContentReport") = HeaderContent
                    Session("BodyReportDs") = ds

                    'Response.Write("<script language=javascript>child=window.open('Preview.aspx');</script>")
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Preview_" & AutoNum, "child=window.open('Preview.aspx');", True)

                Else

                    Dim PesanError As String = ""

                    Dim FileName As String = "LaporanStatusTPKAOK_" & AutoNum
                    Dim FileExt As String = ".csv"
                    Dim FileNameExt As String = FileName & FileExt

                    'Dim hasil As String = ObjFungsi.DataSetToFile(ds, "|", False, FileName, FileExt, "BIGDATA", PesanError)
                    Dim hasil As String = ObjFungsi.DataReaderToFile(MCon, SqlQuery, SqlParam, "|", False, FileName, FileExt, "BIGDATA", PesanError)

                    Dim Result() As String = hasil.Split("|")
                    If Result(0) = "1" Then

                        Dim ZipFileName As String = FileName
                        Dim ZipFileExt As String = ".zip"
                        Dim ZipFileNameExt As String = ZipFileName & ZipFileExt

                        Dim FileToZip As String = "" & Result(1)
                        Dim FilesToZip As String() = FileToZip.Split("|")
                        Dim FileToDownload As String = ObjFungsi.ProsesZip(FilesToZip, "CompressedReport", ZipFileNameExt, PesanError)

                        If FileToDownload <> "" Then

                            Dim fileInfo As FileInfo = New FileInfo(FileToDownload)

                            If fileInfo.Exists Then

                                Session("DownloadPage_FileToDownload") = FileToDownload
                                Session("DownloadPage_DLFileName") = ZipFileName
                                Session("DownloadPage_DLFileExt") = ZipFileExt

                                ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Download_" & AutoNum, "child=window.open('DownloadPage.aspx');", True)

                            Else

                                ObjFungsi.WriteTracelogTxt("File " & FileToDownload & " not exists")
                                lblError.Text = "File " & ZipFileNameExt & " not exists"
                                ScriptManager1.SetFocus(lblError)

                            End If

                        Else

                            ObjFungsi.WriteTracelogTxt("File " & FileNameExt & " fail to create!")
                            lblError.Text = "File " & FileNameExt & " fail to create!"
                            ScriptManager1.SetFocus(lblError)

                        End If

                    Else
                        lblError.Text = Result(1)
                        ScriptManager1.SetFocus(lblError)
                    End If

                End If

            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
                ObjFungsi.WriteTracelogTxt(Pesan)

                lblError.Text = ex.Message
                ScriptManager1.SetFocus(lblError)

            Finally

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

                Try
                    ds = Nothing
                Catch ex As Exception
                End Try

            End Try

        Else

            Response.Redirect("Login.aspx", False)

            Dim VirtualPath As String = Request.CurrentExecutionFilePath
            Dim FileName As String = System.Web.VirtualPathUtility.GetFileName(VirtualPath)
            Session("RedirectMessage") = FileName & ", User Kosong"

        End If

    End Sub

    Private Function GetStatus(ByVal Status As String) As String

        Dim hasil1 As String = ""

        Status = Status.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList1)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Status = "NAME" Then
                        hasil1 = ddlStatus.SelectedItem.Text
                    Else
                        hasil1 = ddlStatus.SelectedValue
                    End If
                Else
                    If Status = "NAME" Then
                        hasil1 = ViewState("Multi_Status")
                    Else
                        hasil1 = ViewState("Multi_Status")
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

    Protected Sub btnMulti1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMulti1.Click, btnMulti2.Click

        lblError.Text = ""

        txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
        txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

        Dim btn As Button = sender
        'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
        Dim btnID As String = (btn.ID).Replace("btnMulti", "")

        If btnID = "1" Then
            td1.Visible = False
            tr1.Visible = True
            ObjFungsi.SetMulti(chkList1, ViewState("Multi_Status"))
        ElseIf btnID = "2" Then
            td2.Visible = False
            tr2.Visible = True
            ObjFungsi.SetMulti(chkList2, ViewState("Multi_Dc"))
        End If

    End Sub

    Protected Sub btnAll1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAll1.Click, btnAll2.Click

        lblError.Text = ""

        Try

            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnAll", "")

            If btnID = "1" Then
                ObjFungsi.SelectAll(chkList1)
            ElseIf btnID = "2" Then
                ObjFungsi.SelectAll(chkList2)
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

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
                        hasil1 = ViewState("Multi_DcName")
                    Else
                        hasil1 = ViewState("Multi_Dc")
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

    Protected Sub btnNon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNon1.Click, btnNon2.Click

        lblError.Text = ""

        Try

            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnNon", "")

            If btnID = "1" Then
                ObjFungsi.DeSelectAll(chkList1)
            ElseIf btnID = "2" Then
                ObjFungsi.DeSelectAll(chkList2)
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Protected Sub btnBatal1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBatal1.Click, btnBatal2.Click

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
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Protected Sub btnPilih1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPilih1.Click, btnPilih2.Click

        lblError.Text = ""

        Try

            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnPilih", "")

            If btnID = "1" Then

                td1.Visible = True
                tr1.Visible = False

                ViewState("Multi_Status") = ObjFungsi.SetDataMulti(btnMulti1, chkList1)
                ViewState("Multi_StatusName") = ObjFungsi.SetDataMultiName(btnMulti1, chkList1)

                ScriptManager1.SetFocus(btnMulti1)
            ElseIf btnID = "2" Then

                td2.Visible = True
                tr2.Visible = False

                ViewState("Multi_Dc") = ObjFungsi.SetDataMulti(btnMulti2, chkList2)
                ViewState("Multi_DcName") = ObjFungsi.SetDataMultiName(btnMulti2, chkList2)

                ScriptManager1.SetFocus(btnMulti2)
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

End Class