
Imports System.Data
Imports ReportWeb.ClsWebVer
Imports MySql.Data.MySqlClient

Partial Class FlagHRB
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

        Try
            ViewState("HasilCek") = "KOSONG"
            ViewState("txtNoAWB") = ""

            Dim dt As DataTable = ObjFungsi.PenyebabDtV2()

            'halaman ini, di-duplikasi menjadi halaman FlagGGL.aspx
            'sehingga di halaman ini exclude status GGL
            dt.DefaultView.RowFilter = "Status <> 'FLR'"
            dt = dt.DefaultView.ToTable

            ObjFungsi.ddlBound(ddlPenyebab, dt, "Display", "Value")
            ViewState("PenyebabDt") = dt

        Catch ex As Exception
            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message
        End Try

    End Sub

    'Private Sub ClearAll()

    '    ddlPenyebab.SelectedIndex = -1
    '    TxtKeterangan.Text = ""
    '    txtNoAWB.Text = ""

    '    gvData.DataSource = Nothing
    '    gvData.DataBind()

    '    LblTotalBaris.Text = "Total Baris : 0"
    '    LblTotalBerhasil.Text = "Total Berhasil : 0"

    '    ViewState("gvData") = Nothing

    'End Sub

    Private Function CreateEmptyDt() As DataTable

        Dim dt As New DataTable

        dt.Columns.Add("TrackNum")
        dt.Columns.Add("Informasi")
        dt.Columns.Add("HasilCek")

        Return dt

    End Function

    Private Sub FillDt(ByVal dtTemp As DataTable, ByVal AWB As String)

        dtTemp.DefaultView.RowFilter = "TrackNum = '" & AWB & "'"
        If dtTemp.DefaultView.Count = 0 Then
            Dim dt As DataTable = CreateEmptyDt()
            dt.Rows.Add()
            dt.Rows(0).Item("TrackNum") = AWB
            dt.Rows(0).Item("Informasi") = ""
            dt.Rows(0).Item("HasilCek") = ""

            Dim dr As DataRow = dt.Rows(0)
            dtTemp.Rows.Add(dr.ItemArray)
        End If

    End Sub

    Private Function ValidasiTambah(ByRef AWBList() As String) As Boolean

        Try

            lblError.Text = ""

            ddlPenyebab.SelectedIndex = -1
            ViewState("old_ddlPenyebab") = ""

            LblTotalBaris.Text = ""

            gvData.DataSource = Nothing
            gvData.DataBind()
            ViewState("gvData") = Nothing

            LblTotalBerhasil.Text = ""

            TrData.Visible = False

            TxtKeterangan.Text = ""

            txtNoAWB.Text = txtNoAWB.Text.Trim

            If txtNoAWB.Text = "" Then
                lblError.Text = lblNoAWB.Text & " masih kosong!"
                Return False
            End If

            If txtNoAWB.Text.Contains(vbLf) Then
                AWBList = txtNoAWB.Text.Split(vbLf)
            Else
                AWBList = txtNoAWB.Text.Split(vbCrLf)
            End If

            If Not IsNothing(AWBList) Then

                For i As Integer = 0 To AWBList.Length - 1
                    AWBList(i) = AWBList(i).Trim
                Next

                If AWBList.Length = 1 And AWBList(0) = "" Then
                    lblError.Text = lblNoAWB.Text & " masih kosong!"
                    Return False
                End If

            Else
                lblError.Text = lblNoAWB.Text & " masih kosong!"
                Return False
            End If

            Return True

        Catch ex As Exception

            lblError.Text = ex.Message
            Return False

        End Try

    End Function

    Protected Sub BtnTambah_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnTambah.Click

        Dim AWBList() As String = Nothing

        If ValidasiTambah(AWBList) Then

            Try

                Dim dtTemp As New DataTable
                'dtTemp = CreateEmptyDt()
                dtTemp.Columns.Add("TrackNum")
                dtTemp.Columns.Add("Informasi")
                dtTemp.Columns.Add("HasilCek")

                Dim drTemp As DataRow

                For Each AWB As String In AWBList

                    If AWB <> "" Then

                        drTemp = dtTemp.NewRow

                        drTemp.Item("TrackNum") = AWB
                        drTemp.Item("Informasi") = ""
                        drTemp.Item("HasilCek") = ""

                        dtTemp.Rows.Add(drTemp)

                    End If

                Next

                'For Each AWB As String In AWBList

                '    If AWB <> "" Then

                '        FillDt(dtTemp, AWB)

                '    End If

                'Next

                'Dim dt As DataTable = ViewState("gvData")

                'If Not IsNothing(dt) Then

                '    For Each dr As DataRow In dtTemp.Rows

                '        dt.DefaultView.RowFilter = "TrackNum = '" & dr.Item("TrackNum") & "'"
                '        If dt.DefaultView.Count = 0 Then
                '            dt.Rows.Add(dr.ItemArray)
                '        End If

                '    Next

                'Else
                '    dt = New DataTable
                '    dt = dtTemp
                'End If

                'dt.DefaultView.RowFilter = Nothing

                ViewState("gvData") = dtTemp

                gvData.DataSource = dtTemp
                gvData.DataBind()

                LblTotalBaris.Text = "Total Baris : " & dtTemp.Rows.Count

                TrData.Visible = True

            Catch ex As Exception

                lblError.Text = ex.Message

            End Try

        End If

    End Sub

    Protected Sub gvData_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs)

        If (e.CommandName.Equals("Hapus")) Then

            Try

                lblError.Text = ""

                Dim IndexRow As Integer = Convert.ToInt32(e.CommandArgument)

                Dim dt As DataTable = ViewState("gvData")
                dt = ObjFungsi.HapusDariDataTable(dt, IndexRow)

                ViewState("dtDraft") = dt

                gvData.DataSource = dt
                gvData.DataBind()

                LblTotalBaris.Text = "Total Baris : " & dt.Rows.Count

                dt.DefaultView.RowFilter = "HasilCek = 'BERHASIL'"
                LblTotalBerhasil.Text = "Total Berhasil : " & dt.DefaultView.Count
                dt.DefaultView.RowFilter = Nothing

            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
                ObjFungsi.WriteTracelogTxt(Pesan)

            End Try

        End If

    End Sub

    Private Function validasiCek() As Boolean

        lblError.Text = ""

        ViewState("old_ddlPenyebab") = ""

        Try

            If ddlPenyebab.SelectedValue = "" Then
                lblError.Text = LblPenyebab.Text & " belum dipilih!"
                Return False
            End If

            Dim dt As DataTable = ViewState("gvData")

            If Not IsNothing(dt) Then
                If dt.Rows.Count = 0 Then
                    lblError.Text = "Daftar AWB masih kosong!"
                    Return False
                End If
            Else
                lblError.Text = "Daftar AWB masih kosong!"
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

    Protected Sub BtnCek_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCek.Click

        Try

            If validasiCek() Then

                Dim dt As DataTable = ViewState("gvData")

                Dim TrackNum As String = ""
                Dim Informasi As String = ""
                Dim HasilCek As String = ""

                Dim Penyebab As String = ddlPenyebab.SelectedValue
                Dim Status As String = ""

                If Penyebab = "" Then

                Else
                    Status = "OOC" 'secara umum akan CheckTransaction OOC

                    Dim dtPenyebab As DataTable = ViewState("PenyebabDt")
                    dtPenyebab.DefaultView.RowFilter = "Value = '" & Penyebab & "'"
                    If dtPenyebab.DefaultView.Count > 0 Then
                        If dtPenyebab.DefaultView(0).Item("Status").ToString <> "" _
                        And dtPenyebab.DefaultView(0).Item("Status").ToString.Trim.ToUpper <> Status Then
                            Status = dtPenyebab.DefaultView(0).Item("Status").ToString.Trim.ToUpper
                        End If
                    End If
                End If

                If Status <> "" Then
                    For i As Integer = 0 To dt.Rows.Count - 1
                        Informasi = ""
                        HasilCek = ""

                        CheckTransaction(dt.Rows(i).Item("TrackNum"), Status, Informasi, HasilCek)

                        dt.Rows(i).Item("Informasi") = Informasi
                        dt.Rows(i).Item("HasilCek") = HasilCek
                    Next
                End If

                ViewState("old_ddlPenyebab") = ddlPenyebab.SelectedValue

                ViewState("gvData") = dt

                gvData.DataSource = dt
                gvData.DataBind()

                LblTotalBaris.Text = "Total Baris : " & dt.Rows.Count

                dt.DefaultView.RowFilter = "HasilCek = 'BERHASIL'"
                LblTotalBerhasil.Text = "Total Berhasil : " & dt.DefaultView.Count
                dt.DefaultView.RowFilter = Nothing

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Function Validasi(ByRef AWBList As DataTable) As Boolean

        lblError.Text = ""

        'Validasi Penyebab
        Dim old_ddlPenyebab As String = "" & ViewState("old_ddlPenyebab")

        If ddlPenyebab.SelectedValue <> old_ddlPenyebab Then
            lblError.Text = LblPenyebab.Text & " yang dipilih berubah, silahkan tekan tombol " & BtnCek.Text & " kembali!"
            Return False
        End If

        If ddlPenyebab.SelectedValue = "" Then
            lblError.Text = LblPenyebab.Text & " belum dipilih!"
            Return False
        End If


        'Validasi Keterangan
        TxtKeterangan.Text = TxtKeterangan.Text.Trim
        If TxtKeterangan.Text = "" Then
            lblError.Text = LblKeterangan.Text & " belum diisi!"
            Return False
        End If


        'Validasi Daftar AWB
        Dim dt As DataTable = ViewState("gvData")

        If Not IsNothing(dt) Then
            If dt.Rows.Count = 0 Then
                lblError.Text = "Daftar AWB masih kosong!"
                Return False
            End If
        Else
            lblError.Text = "Daftar AWB masih kosong!"
            Return False
        End If

        dt.DefaultView.RowFilter = "HasilCek = 'BERHASIL'"

        If dt.DefaultView.Count = 0 Then
            lblError.Text = "Hasil Cek Daftar tidak ada yang BERHASIL!"
            Return False
        Else
            AWBList = dt.DefaultView.ToTable(True, "TrackNum")
            AWBList.Columns.Add("StatusUpdate")
        End If

        Return True

    End Function

    Protected Sub BtnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnProses.Click

        Try

            Dim AWBList As New DataTable
            Dim Result As String = ""
            Dim SB As New StringBuilder
            Dim SBResponUpdateRTSDate As New StringBuilder

            If Validasi(AWBList) Then

                Dim Penyebab As String = ddlPenyebab.SelectedValue
                Dim Status As String = ""

                If Penyebab = "" Then

                Else
                    Status = "OOC" 'secara umum akan CheckTransaction OOC

                    Dim dtPenyebab As DataTable = ViewState("PenyebabDt")
                    dtPenyebab.DefaultView.RowFilter = "Value = '" & Penyebab & "'"
                    If dtPenyebab.DefaultView.Count > 0 Then
                        If dtPenyebab.DefaultView(0).Item("Status").ToString <> "" _
                        And dtPenyebab.DefaultView(0).Item("Status").ToString.Trim.ToUpper <> Status Then
                            Status = dtPenyebab.DefaultView(0).Item("Status").ToString.Trim.ToUpper
                        End If
                    End If
                End If

                If Status <> "" Then

                    For i As Integer = 0 To AWBList.Rows.Count - 1

                        If i > 0 Then
                            SB.Append("<br/>")
                        End If

                        If CustomUpdateStatus(AWBList.Rows(i).Item("TrackNum"), Status, Result) Then
                            AWBList.Rows(i).Item("StatusUpdate") = "1"
                        Else
                            AWBList.Rows(i).Item("StatusUpdate") = "0"
                        End If

                        Result &= " - " & AWBList.Rows(i).Item("TrackNum")

                        SB.Append(Result)

                    Next

                End If

                '2021-12-20 15:14, By Cucun, Req Pa Paulus, Tambah Proses Update RTSDate, dan Kirim Email, Untuk Status XDO
                If Status = "XDO" Then

                    AWBList.DefaultView.RowFilter = "StatusUpdate = '1'"

                    If AWBList.DefaultView.Count > 0 Then

                        Dim SBAWB As New StringBuilder

                        For i As Integer = 0 To AWBList.DefaultView.Count - 1

                            If i > 0 Then
                                SBAWB.Append(",")
                            End If

                            SBAWB.Append("'" & AWBList.DefaultView(i).Item("TrackNum") & "'")

                        Next

                        AWBList.DefaultView.RowFilter = Nothing

                        Dim ListAWB As String = SBAWB.ToString

                        Dim PesanError As String = ""

                        Dim dt As DataTable = ObjFungsi.UpdateRTSDate_getDaftar(ListAWB, True, PesanError)

                        If Not IsNothing(dt) Then
                            If dt.Rows.Count > 0 Then

                                Dim UserID As String = "" & Session("NIK")

                                If UserID <> "" Then

                                    Dim ErrorMessage As String = ""

                                    If ObjFungsi.UpdateRTSDate_SimpanData(dt, UserID, ErrorMessage) Then

                                        ErrorMessage = ""

                                        If ObjFungsi.UpdateRTSDate_KirimEmail(dt, ErrorMessage) Then
                                            SBResponUpdateRTSDate.Append("UpdateRTSDate Simpan Data dan Kirim Email Berhasil!")
                                        Else
                                            SBResponUpdateRTSDate.Append("UpdateRTSDate Gagal Kirim Email!")
                                            If ErrorMessage <> "" Then
                                                SBResponUpdateRTSDate.Append(" - Error : " & ErrorMessage)
                                            End If
                                        End If

                                    Else
                                        SBResponUpdateRTSDate.Append("UpdateRTSDate Gagal Simpan Data!")
                                        If ErrorMessage <> "" Then
                                            SBResponUpdateRTSDate.Append(" - Error : " & ErrorMessage)
                                        End If
                                    End If

                                End If

                            Else
                                SBResponUpdateRTSDate.Append("UpdateRTSDate_getDaftar, Tidak ada data!")
                                If PesanError <> "" Then
                                    SBResponUpdateRTSDate.Append(" - Error : " & PesanError)
                                End If
                            End If
                        Else
                            SBResponUpdateRTSDate.Append("getDaftar Nothing!")
                            If PesanError <> "" Then
                                SBResponUpdateRTSDate.Append(" - Error : " & PesanError)
                            End If
                        End If

                    End If

                End If

                Dim ResponUpdateRTSDate As String = SBResponUpdateRTSDate.ToString

                If ResponUpdateRTSDate <> "" Then
                    SB.Append("<br/><br/>")
                    SB.Append(ResponUpdateRTSDate)
                End If

                Session("ResultMessage" & MyPage) = SB.ToString
                Response.Redirect(Request.RawUrl, False)

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Sub CheckTransaction(ByVal Tracknum As String, ByVal Status As String, ByRef Informasi As String, ByRef HasilCek As String)

        Try

            ReDim param(7)
            param(0) = UserWS
            param(1) = PassWS
            param(2) = Status 'STATUS
            param(3) = "IPP" 'COMPANY
            param(4) = "RPT" 'HUB/STA/...
            param(5) = Session("NIK")
            param(6) = Tracknum 'SINGLE TRACKNUM
            param(7) = ""

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.CheckTransaction(AppName, AppVersion, param)

            If respon(0) = "0" Then

                Dim dt As DataTable = ObjFungsi.ConvertStringToDatatable(respon(2).ToString)

                Informasi = "Partner : " & dt.Rows(0).Item(1).ToString
                Informasi &= "<br/>Penerima : " & dt.Rows(0).Item(2).ToString
                Informasi &= "<br/>Last Status : " & dt.Rows(0).Item(3).ToString & " / " & dt.Rows(0).Item(4).ToString

                HasilCek = "BERHASIL"

            Else
                Informasi = "GAGAL - " & respon(1)
                HasilCek = "GAGAL"
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            Informasi = "GAGAL - " & ex.Message

            HasilCek = "GAGAL"

        End Try

    End Sub

    Private Function CustomUpdateStatus(ByVal Tracknum As String, ByVal Status As String, ByRef Result As String) As Boolean

        Try
            Dim Penyebab As String = ddlPenyebab.SelectedValue

            ReDim param(7)
            param(0) = UserWS
            param(1) = PassWS
            param(2) = Status
            param(3) = "RPT"
            param(4) = Session("NIK")
            param(5) = Session("NameOfUser")
            param(6) = Tracknum
            param(7) = Penyebab & " - " & TxtKeterangan.Text.Replace(vbCrLf, " ").Replace("  ", " ")

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.CustomUpdateStatus(AppName, AppVersion, param)

            If respon(0) = "0" Then
                Result = "Berhasil !!"
                Return True
            Else
                Result = respon(1)
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            Result = ex.Message

        End Try

        Return False

    End Function

End Class
