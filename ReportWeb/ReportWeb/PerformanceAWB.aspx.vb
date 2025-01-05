
Imports System.Data
Imports ReportWeb.ClsWebVer
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Collections.Generic
Imports Microsoft.VisualBasic.ApplicationServices

Partial Class PerformanceAWB
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

            Dim dt As DataTable = GetRegionList()
            ObjFungsi.ddlBound(ddlRegionAsal, dt, "Display", "Value")
            ObjFungsi.ddlBound(ddlRegion, dt, "Display", "Value")

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList1, dt, "Display", "Value")
            ObjFungsi.chkBound(chkList3, dt, "Display", "Value")

            dt = GetHubList()
            ObjFungsi.ddlBound(ddlAsal, dt, "Display", "Value")
            ObjFungsi.ddlBound(ddlTujuan, dt, "Display", "Value")

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList2, dt, "Display", "Value")
            ObjFungsi.chkBound(chkList4, dt, "Display", "Value")

            dt = ObjFungsi.GetServiceTypeList(Me.GetType().Name, "")
            ObjFungsi.ddlBound(ddlLayanan, dt, "Display", "Value")

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList40, dt, "Display", "Value")

            dt = ObjFungsi.GetECommerceList(Me.GetType().Name, Session("UserCode"))
            ObjFungsi.ddlBound(ddlECom, dt, "Display", "Value")

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList30, dt, "Display", "Value")

            'dt = ObjFungsi.GetAccountCategory(ErrorMessage)
            'ObjFungsi.ddlBound(ddlTipeSeller, dt, "Display", "Value")

            dt = ObjFungsi.GetAccountCategory(ErrorMessage)
            ObjFungsi.ddlBound(ddlKategoriSeller, dt, "Display", "Value")

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList6, dt, "Display", "Value")

            dt = ObjFungsi.GetAccountSubCategory(ErrorMessage)
            ObjFungsi.ddlBound(ddlSubKategoriSeller, dt, "Display", "Value")

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList7, dt, "Display", "Value")

            dt = ObjFungsi.ConvertStringToDatatable("Display,Value|,|SELESAI,SELESAI|BELUM SELESAI,BELUM SELESAI")
            ObjFungsi.ddlBound(ddlKategori, dt, "Display", "Value")


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

    Protected Sub btnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProses.Click

        ProsesPreview(False)

    End Sub

    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreview.Click

        ProsesPreview(True)

    End Sub

    Private Function ValidasiProses() As Boolean

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

            Return False

        End Try

    End Function

    Private Sub ProsesPreview(ByVal PreviewOnly As Boolean)

        Try

            If ValidasiProses() Then

                Dim RegionAsal As String = GetRegionAsal("CODE")
                Dim RegionAsalName As String = GetRegionAsal("NAME")
                Dim HubAsal As String = GetHubAsal("CODE")
                Dim HubAsalName As String = GetHubAsal("NAME")
                Dim RegionTujuan As String = GetRegion("CODE")
                Dim RegionTujuanName As String = GetRegion("NAME")
                Dim HubTujuan As String = GetHub("CODE")
                Dim HubTujuanName As String = GetHub("NAME")
                Dim KodeEcommerce As String = GetEcommerce("CODE")
                Dim NamaEcommerce As String = GetEcommerce("NAME")
                'Dim KodeTipeSeller As String = ddlTipeSeller.SelectedValue
                'Dim NamaTipeSeller As String = ddlTipeSeller.SelectedItem.Text
                Dim KategoriSeller As String = GetKategoriSeller("CODE")
                Dim NamaKategoriSeller As String = GetKategoriSeller("NAME")
                Dim SubKategoriSeller As String = GetSubKategoriSeller("CODE")
                Dim NamaSubKategoriSeller As String = GetSubKategoriSeller("NAME")
                Dim TanggalAwal As String = txtTgl1.Text
                Dim TanggalAkhir As String = txtTgl2.Text
                Dim AsFinalDate As String = ChkAsFinalDate.Checked.ToString.ToUpper
                Dim KodeLayanan As String = GetJenisLayanan("CODE")
                Dim Kategori As String = ddlKategori.SelectedValue
                Dim NamaLayanan As String = GetJenisLayanan("NAME")

                Proses(RegionAsal, RegionAsalName, HubAsal, HubAsalName, RegionTujuan, RegionTujuanName, HubTujuan, HubTujuanName, KodeEcommerce, NamaEcommerce, KategoriSeller, NamaKategoriSeller, SubKategoriSeller, NamaSubKategoriSeller, TanggalAwal, TanggalAkhir, AsFinalDate, KodeLayanan, NamaLayanan, Kategori, PreviewOnly)

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Sub Proses(ByVal RegionAsal As String, ByVal RegionAsalName As String, ByVal HubAsal As String, ByVal HubAsalName As String, ByVal RegionTujuan As String, ByVal RegionTujuanName As String, ByVal HubTujuan As String, ByVal HubTujuanName As String, ByVal KodeEcommerce As String, ByVal NamaEcommerce As String, ByVal KategoriSeller As String, ByVal NamaKategoriSeller As String, ByVal SubKategoriSeller As String, ByVal NamaSubKategoriSeller As String, ByVal TanggalAwal As String, ByVal TanggalAkhir As String, ByVal AsFinalDate As String, ByVal KodeLayanan As String, ByVal NamaLayanan As String, ByVal Kategori As String, ByVal PreviewOnly As Boolean)

        Dim User As String = "" & Session("NIK")

        If User <> "" Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing
            Dim ds As DataSet = Nothing

            Try

                SqlQuery = "call `report_performance_awb`(@RegionAsal,@HubAsal,@RegionTujuan,@HubTujuan,@KodeEcommerce,@KategoriSeller,@SubKategoriSeller,@TanggalAwal,@TanggalAkhir,@AsFinalDate,@KodeLayanan, @Kategori, @PreviewOnly);"

                Dim SqlParam As New Dictionary(Of String, String)
                SqlParam.Add("@RegionAsal", RegionAsal)
                SqlParam.Add("@HubAsal", HubAsal)
                SqlParam.Add("@RegionTujuan", RegionTujuan)
                SqlParam.Add("@HubTujuan", HubTujuan)
                SqlParam.Add("@KodeEcommerce", KodeEcommerce)
                SqlParam.Add("@KategoriSeller", KategoriSeller)
                SqlParam.Add("@SubKategoriSeller", SubKategoriSeller)
                SqlParam.Add("@TanggalAwal", TanggalAwal)
                SqlParam.Add("@TanggalAkhir", TanggalAkhir)
                SqlParam.Add("@AsFinalDate", AsFinalDate)
                SqlParam.Add("@KodeLayanan", KodeLayanan)
                SqlParam.Add("@Kategori", Kategori)
                SqlParam.Add("@PreviewOnly", PreviewOnly)

                MCon = ObjCon.SetConn_Slave1
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

                If PreviewOnly Then

                    Dim HeaderRow As Integer = 11
                    Dim HeaderTitle(HeaderRow) As String
                    Dim HeaderContent(HeaderRow) As String

                    HeaderTitle(0) = "REGION(HUB ASAL)"
                    HeaderTitle(1) = "HUB ASAL"
                    HeaderTitle(2) = "REGION(HUB TUJUAN)"
                    HeaderTitle(3) = "HUB TUJUAN"
                    HeaderTitle(4) = "KATEGORI SELLER"
                    HeaderTitle(5) = "SUB KATEGORI SELLER"
                    HeaderTitle(6) = "NAMA E-COMM"
                    HeaderTitle(7) = "TANGGAL AWAL"
                    HeaderTitle(8) = "TANGGAL AKHIR"
                    HeaderTitle(9) = "SEBAGAI TGL.FINAL"
                    HeaderTitle(10) = "JENIS LAYANAN"
                    HeaderTitle(11) = "KATEGORI"

                    HeaderContent(0) = RegionAsalName
                    HeaderContent(1) = HubAsalName
                    HeaderContent(2) = RegionTujuanName
                    HeaderContent(3) = HubTujuanName
                    HeaderContent(4) = NamaKategoriSeller
                    HeaderContent(5) = NamaSubKategoriSeller
                    HeaderContent(6) = NamaEcommerce
                    HeaderContent(7) = TanggalAwal
                    HeaderContent(8) = TanggalAkhir
                    HeaderContent(9) = IIf(AsFinalDate = "1" Or AsFinalDate = "TRUE", "YA", "TIDAK")
                    HeaderContent(10) = NamaLayanan
                    HeaderContent(11) = Kategori

                    ds = ObjSQL.SQLInsertIntoDataset(MCon, SqlQuery, SqlParam)

                    Session("TitleReport") = "PERFORMANCE AWB"
                    Session("HeaderTitleReport") = HeaderTitle
                    Session("HeaderContentReport") = HeaderContent
                    Session("BodyReportDs") = ds

                    'Response.Write("<script language=javascript>child=window.open('Preview.aspx');</script>")
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Preview_" & AutoNum, "child=window.open('Preview.aspx');", True)

                Else

                    Dim PesanError As String = ""

                    Dim FileName As String = "PerformanceAWB_" & AutoNum
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

#Region "Button Multi Select"

    Private Function GetRegion(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList1)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlRegion.SelectedItem.Text
                    Else
                        hasil1 = ddlRegion.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_RegionName")
                    Else
                        hasil1 = ViewState("Multi_Region")
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

    Private Function GetRegionAsal(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList3)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlRegionAsal.SelectedItem.Text
                    Else
                        hasil1 = ddlRegionAsal.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_RegionAsalName")
                    Else
                        hasil1 = ViewState("Multi_RegionAsal")
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

    Private Function GetHub(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList2)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlTujuan.SelectedItem.Text
                    Else
                        hasil1 = ddlTujuan.SelectedValue
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

    Private Function GetHubAsal(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList4)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlAsal.SelectedItem.Text
                    Else
                        hasil1 = ddlAsal.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_HubAsalName")
                    Else
                        hasil1 = ViewState("Multi_HubAsal")
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

    Private Function GetEcommerce(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList30)

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

    Private Function GetKategoriSeller(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList6)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlKategoriSeller.SelectedItem.Text
                    Else
                        hasil1 = ddlKategoriSeller.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_KategoriSellerName")
                    Else
                        hasil1 = ViewState("Multi_KategoriSeller")
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

    Private Function GetSubKategoriSeller(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList7)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlSubKategoriSeller.SelectedItem.Text
                    Else
                        hasil1 = ddlSubKategoriSeller.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_SubKategoriSellerName")
                    Else
                        hasil1 = ViewState("Multi_SubKategoriSeller")
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

    Private Function GetJenisLayanan(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList40)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlLayanan.SelectedItem.Text
                    Else
                        hasil1 = ddlLayanan.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_JenisLayananName")
                    Else
                        hasil1 = ViewState("Multi_JenisLayanan")
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

    Protected Sub btnMulti1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMulti1.Click, btnMulti2.Click, btnMulti3.Click, btnMulti4.Click, btnMulti30.Click, btnMulti40.Click, btnMulti6.Click, btnMulti7.Click

        lblError.Text = ""

        txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
        txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

        Dim btn As Button = sender
        'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
        Dim btnID As String = (btn.ID).Replace("btnMulti", "")

        If btnID = "1" Then
            td1.Visible = False
            tr1.Visible = True
            ObjFungsi.SetMulti(chkList1, ViewState("Multi_Region"))
        ElseIf btnID = "2" Then
            td2.Visible = False
            tr2.Visible = True
            ObjFungsi.SetMulti(chkList2, ViewState("Multi_Hub"))
        ElseIf btnID = "3" Then
            td3.Visible = False
            tr3.Visible = True
            ObjFungsi.SetMulti(chkList3, ViewState("Multi_RegionAsal"))
        ElseIf btnID = "4" Then
            td4.Visible = False
            tr4.Visible = True
            ObjFungsi.SetMulti(chkList4, ViewState("Multi_HubAsal"))
        ElseIf btnID = "30" Then
            td30.Visible = False
            tr30.Visible = True
            ObjFungsi.SetMulti(chkList30, ViewState("Multi_Ecomm"))
        ElseIf btnID = "40" Then
            td40.Visible = False
            tr40.Visible = True
            ObjFungsi.SetMulti(chkList40, ViewState("Multi_JenisLayanan"))
        ElseIf btnID = "6" Then
            td6.Visible = False
            tr6.Visible = True
            ObjFungsi.SetMulti(chkList6, ViewState("Multi_KategoriSeller"))
        ElseIf btnID = "7" Then
            td7.Visible = False
            tr7.Visible = True
            ObjFungsi.SetMulti(chkList7, ViewState("Multi_SubKategoriSeller"))
        Else
        End If

    End Sub

    Protected Sub btnAll1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAll1.Click, btnAll2.Click, btnAll3.Click, btnAll4.Click, btnAll30.Click, btnAll40.Click, btnAll6.Click, btnAll7.Click

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
            ElseIf btnID = "4" Then
                ObjFungsi.SelectAll(chkList4)
            ElseIf btnID = "30" Then
                ObjFungsi.SelectAll(chkList30)
            ElseIf btnID = "40" Then
                ObjFungsi.SelectAll(chkList40)
            ElseIf btnID = "6" Then
                ObjFungsi.SelectAll(chkList6)
            ElseIf btnID = "7" Then
                ObjFungsi.SelectAll(chkList7)
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

    Protected Sub btnNon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNon1.Click, btnNon2.Click, btnNon3.Click, btnNon4.Click, btnNon30.Click, btnNon40.Click, btnNon6.Click, btnNon7.Click

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
            ElseIf btnID = "4" Then
                ObjFungsi.DeSelectAll(chkList4)
            ElseIf btnID = "30" Then
                ObjFungsi.DeSelectAll(chkList30)
            ElseIf btnID = "40" Then
                ObjFungsi.DeSelectAll(chkList40)
            ElseIf btnID = "6" Then
                ObjFungsi.DeSelectAll(chkList6)
            ElseIf btnID = "7" Then
                ObjFungsi.DeSelectAll(chkList7)
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

    Protected Sub btnBatal1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBatal1.Click, btnBatal2.Click, btnBatal3.Click, btnBatal4.Click, btnBatal30.Click, btnBatal40.Click, btnBatal6.Click, btnBatal7.Click

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
            ElseIf btnID = "4" Then
                td4.Visible = True
                tr4.Visible = False
                ScriptManager1.SetFocus(btnMulti4)
            ElseIf btnID = "30" Then
                td30.Visible = True
                tr30.Visible = False
                ScriptManager1.SetFocus(btnMulti30)
            ElseIf btnID = "40" Then
                td40.Visible = True
                tr40.Visible = False
                ScriptManager1.SetFocus(btnMulti40)
            ElseIf btnID = "6" Then
                td6.Visible = True
                tr6.Visible = False
                ScriptManager1.SetFocus(btnMulti6)
            ElseIf btnID = "7" Then
                td7.Visible = True
                tr7.Visible = False
                ScriptManager1.SetFocus(btnMulti7)
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

    Protected Sub btnPilih1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPilih1.Click, btnPilih2.Click, btnPilih3.Click, btnPilih4.Click, btnPilih30.Click, btnPilih40.Click, btnPilih6.Click, btnPilih7.Click

        lblError.Text = ""

        Try

            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnPilih", "")

            If btnID = "1" Then

                td1.Visible = True
                tr1.Visible = False

                ViewState("Multi_Region") = ObjFungsi.SetDataMulti(btnMulti1, chkList1)
                ViewState("Multi_RegionName") = ObjFungsi.SetDataMultiName(btnMulti1, chkList1)

                ScriptManager1.SetFocus(btnMulti1)

            ElseIf btnID = "2" Then

                td2.Visible = True
                tr2.Visible = False

                ViewState("Multi_Hub") = ObjFungsi.SetDataMulti(btnMulti2, chkList2)
                ViewState("Multi_HubName") = ObjFungsi.SetDataMultiName(btnMulti2, chkList2)

                ScriptManager1.SetFocus(btnMulti2)

            ElseIf btnID = "3" Then

                td3.Visible = True
                tr3.Visible = False

                ViewState("Multi_RegionAsal") = ObjFungsi.SetDataMulti(btnMulti3, chkList3)
                ViewState("Multi_RegionAsalName") = ObjFungsi.SetDataMultiName(btnMulti3, chkList3)

                ScriptManager1.SetFocus(btnMulti3)

            ElseIf btnID = "4" Then

                td4.Visible = True
                tr4.Visible = False

                ViewState("Multi_HubAsal") = ObjFungsi.SetDataMulti(btnMulti4, chkList4)
                ViewState("Multi_HubAsalName") = ObjFungsi.SetDataMultiName(btnMulti4, chkList4)

                ScriptManager1.SetFocus(btnMulti4)

            ElseIf btnID = "30" Then

                td30.Visible = True
                tr30.Visible = False

                ViewState("Multi_Ecomm") = ObjFungsi.SetDataMulti(btnMulti30, chkList30)
                ViewState("Multi_EcommName") = ObjFungsi.SetDataMultiName(btnMulti30, chkList30)

                ScriptManager1.SetFocus(btnMulti30)

            ElseIf btnID = "40" Then

                td40.Visible = True
                tr40.Visible = False

                ViewState("Multi_JenisLayanan") = ObjFungsi.SetDataMulti(btnMulti40, chkList40)
                ViewState("Multi_JenisLayananName") = ObjFungsi.SetDataMultiName(btnMulti40, chkList40)

                ScriptManager1.SetFocus(btnMulti40)

            ElseIf btnID = "6" Then

                td6.Visible = True
                tr6.Visible = False

                ViewState("Multi_KategoriSeller") = ObjFungsi.SetDataMulti(btnMulti6, chkList6)
                ViewState("Multi_KategoriSellerName") = ObjFungsi.SetDataMultiName(btnMulti6, chkList6)

                ScriptManager1.SetFocus(btnMulti6)

            ElseIf btnID = "7" Then

                td7.Visible = True
                tr7.Visible = False

                ViewState("Multi_SubKategoriSeller") = ObjFungsi.SetDataMulti(btnMulti7, chkList7)
                ViewState("Multi_SubKategoriSellerName") = ObjFungsi.SetDataMultiName(btnMulti7, chkList7)

                ScriptManager1.SetFocus(btnMulti7)

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

#Region "webservice"

    Private Function GetRegionList() As DataTable

        Dim dt As DataTable = Nothing

        dt = ObjFungsi.GetRegionList(Me.GetType().Name)

        If dt.TableName = "ERROR" Then
            lblError.Text = dt.Rows(0).Item("RESPON")
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

    Private Function GetTrackingStatusList() As DataTable

        Dim dt As DataTable = Nothing

        Try
            ReDim param(1)
            param(0) = UserWS
            param(1) = PassWS

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetTrackingStatusListForReport(AppName, AppVersion, param)

            If respon(0) = "0" Then

                dt = ObjFungsi.ConvertStringToDatatable(respon(2))
                dt = dt.DefaultView.ToTable

            Else

                lblError.Text = respon(1)

            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

            Return Nothing

        End Try

        Return dt

    End Function

    Private Function GetTrackingStatusListSimple() As DataTable

        Dim dt As DataTable = Nothing

        Try

            dt = New DataTable

            dt.Columns.Add("Value", GetType(String))
            dt.Columns.Add("Display", GetType(String))

            dt.Rows.Add(New Object() {"", ""})
            dt.Rows.Add(New Object() {"TOKO AWAL", "TOKO AWAL"})
            dt.Rows.Add(New Object() {"HUB", "HUB"})
            dt.Rows.Add(New Object() {"TOKO", "TOKO"})
            dt.Rows.Add(New Object() {"KONSUMEN", "KONSUMEN"})
            dt.Rows.Add(New Object() {"RET", "RET"})
            dt.Rows.Add(New Object() {"GGL", "GGL"})
            dt.Rows.Add(New Object() {"HRB", "HRB"})

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

            Return Nothing

        End Try

        Return dt

    End Function

#End Region

End Class
