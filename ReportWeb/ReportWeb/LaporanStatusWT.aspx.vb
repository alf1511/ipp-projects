Imports System.Data
Imports ReportWeb.ClsWebVer
Imports System.IO
Imports System.Collections.Generic
Imports MySql.Data.MySqlClient

Partial Class LaporanStatusWT
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
        dt = ObjFungsi.GetExpeditionList(Me.GetType().Name, Session("UserCode"))
        ObjFungsi.ddlBound(ddlEkspedisi, dt, "Display", "Value")
        If dt.TableName.ToUpper <> "ERROR" Then
            ObjFungsi.ddlBound(ddlEkspedisi, dt, "Display", "Value")

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList1, dt, "Display", "Value")
        Else
            ErrorList.AppendLine("ERROR GetExpeditionList!")
        End If

        dt.TableName = ""
        dt = ObjFungsi.GetExpeditionExpenseTypeList(Me.GetType().Name, Session("UserCode"), True)
        If dt.TableName.ToUpper <> "ERROR" Then
            ObjFungsi.ddlBound(ddlExpenseType, dt, "Display", "Value")

            dt.DefaultView.RowFilter = "Value <> ''"
            dt = dt.DefaultView.ToTable
            ObjFungsi.chkBound(chkList2, dt, "Display", "Value")
        Else
            ErrorList.AppendLine("ERROR GetExpeditionExpenseTypeList!")
        End If
    End Sub

    Protected Sub btnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProses.Click

        Proses_Preview(False)

    End Sub

    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreview.Click

        Proses_Preview(True)

    End Sub

    Private Function ValidasiProses(ByRef NoAwb As String) As Boolean
        Try

            lblError.Text = ""

            txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
            txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

            Dim ask_date As Boolean = True

            NoAwb = TxtNoAwb.Text.Trim.Replace(vbCrLf, "|")
            Dim Awb() As String = NoAwb.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)

            If Awb.Length > 100 Then
                lblError.Text = "Maksimal Jumlah Nomor Resi 100 !!"
                Return False
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

    Private Sub Proses_Preview(ByVal PreviewOnly As Boolean)
        Dim NoAwb As String = ""
        Try
            If ValidasiProses(NoAwb) Then

                Dim TanggalAwal As String = txtTgl1.Text
                Dim TanggalAkhir As String = txtTgl2.Text
                Dim KodeEkspedisi As String = GetEkspedisi3PL("CODE")
                Dim NamaEkspedisi As String = GetEkspedisi3PL("NAME")
                Dim ExpenseType As String = GetExpenseType("CODE")
                Dim ExpenseTypeDesc As String = GetExpenseType("NAME")

                Proses(TanggalAwal, TanggalAkhir, KodeEkspedisi, NamaEkspedisi, ExpenseType, ExpenseTypeDesc, NoAwb, PreviewOnly)

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Sub Proses(ByVal TanggalAwal As String, ByVal TanggalAkhir As String, ByVal KodeEkspedisi As String, ByVal NamaEkspedisi As String, ByVal ExpenseType As String, ByVal ExpenseTypeDesc As String, ByVal NoAwb As String, ByVal PreviewOnly As Boolean)

        Dim User As String = "" & Session("NIK")

        If User <> "" Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing
            Dim ds As DataSet = Nothing

            Try

                SqlQuery = "call `report_generate_status_wt`(@TanggalAwal,@TanggalAkhir,@KodeEkspedisi,@Proses,@NoAwb,@PreviewOnly);"

                Dim SqlParam As New Dictionary(Of String, String)
                SqlParam.Add("@KodeEkspedisi", KodeEkspedisi)
                SqlParam.Add("@TanggalAwal", TanggalAwal)
                SqlParam.Add("@TanggalAkhir", TanggalAkhir)
                SqlParam.Add("@NoAwb", NoAwb)
                SqlParam.Add("@Proses", ExpenseType)
                SqlParam.Add("@PreviewOnly", PreviewOnly.ToString.ToUpper)

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
                    HeaderTitle(2) = "EKSPEDISI"
                    HeaderTitle(3) = "TIPE PROSES"
                    HeaderTitle(4) = "NOMOR AWB"

                    HeaderContent(0) = TanggalAwal
                    HeaderContent(1) = TanggalAkhir
                    HeaderContent(2) = NamaEkspedisi
                    HeaderContent(3) = ExpenseTypeDesc
                    HeaderContent(4) = NoAwb

                    ds = ObjSQL.SQLInsertIntoDataset(MCon, SqlQuery, SqlParam)

                    Session("TitleReport") = "LAPORAN GENERATE STATUS WT"
                    Session("HeaderTitleReport") = HeaderTitle
                    Session("HeaderContentReport") = HeaderContent
                    Session("BodyReportDs") = ds

                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Preview_" & AutoNum, "child=window.open('Preview.aspx');", True)

                Else

                    Dim PesanError As String = ""

                    Dim FileName As String = "ReportStatusWT_" & AutoNum
                    Dim FileExt As String = ".csv"
                    Dim FileNameExt As String = FileName & FileExt

                    Dim hasil As String = ObjFungsi.DataReaderToFile(MCon, SqlQuery, SqlParam, "|", False, FileName, FileExt, "BIGDATA", PesanError)

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

#Region "Button MultiSelect"

    Private Function GetEkspedisi3PL(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList1)

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

    Private Function GetExpenseType(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList2)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlExpenseType.SelectedItem.Text
                    Else
                        hasil1 = ddlExpenseType.SelectedValue
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
            Return Nothing

        End Try

        Return hasil1

    End Function

    Protected Sub btnMulti1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMulti1.Click, btnMulti2.Click

        lblError.Text = ""

        txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
        txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

        Dim btn As Button = sender
        Dim btnID As String = (btn.ID).Replace("btnMulti", "")

        If btnID = "1" Then
            td1.Visible = False
            tr1.Visible = True
            ObjFungsi.SetMulti(chkList1, ViewState("Multi_Ekspedisi"))
        ElseIf btnID = "2" Then
            td2.Visible = False
            tr2.Visible = True
            ObjFungsi.SetMulti(chkList2, ViewState("Multi_Proses"))
        Else
        End If

    End Sub

    Protected Sub btnAll1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAll1.Click, btnAll2.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            Dim btnID As String = (btn.ID).Replace("btnAll", "")

            If btnID = "1" Then
                ObjFungsi.SelectAll(chkList1)
            ElseIf btnID = "2" Then
                ObjFungsi.SelectAll(chkList2)
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

    Protected Sub btnNon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNon1.Click, btnNon2.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            Dim btnID As String = (btn.ID).Replace("btnNon", "")

            If btnID = "1" Then
                ObjFungsi.DeSelectAll(chkList1)
            ElseIf btnID = "2" Then
                ObjFungsi.DeSelectAll(chkList2)
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

    Protected Sub btnBatal1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBatal1.Click, btnBatal2.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            Dim btnID As String = (btn.ID).Replace("btnBatal", "")

            If btnID = "1" Then
                td1.Visible = True
                tr1.Visible = False
                ScriptManager1.SetFocus(btnMulti1)
            ElseIf btnID = "2" Then
                td2.Visible = True
                tr2.Visible = False
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

    Protected Sub btnPilih1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPilih1.Click, btnPilih2.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            Dim btnID As String = (btn.ID).Replace("btnPilih", "")

            If btnID = "1" Then

                td1.Visible = True
                tr1.Visible = False

                ViewState("Multi_Ekspedisi") = ObjFungsi.SetDataMulti(btnMulti1, chkList1)
                ViewState("Multi_EkspedisiName") = ObjFungsi.SetDataMultiName(btnMulti1, chkList1)

                ScriptManager1.SetFocus(btnMulti1)

            ElseIf btnID = "2" Then

                td2.Visible = True
                tr2.Visible = False

                ViewState("Multi_Proses") = ObjFungsi.SetDataMulti(btnMulti2, chkList2)
                ViewState("Multi_ProsesName") = ObjFungsi.SetDataMultiName(btnMulti2, chkList2)

                ScriptManager1.SetFocus(btnMulti1)

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
