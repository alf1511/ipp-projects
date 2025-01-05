
Imports System.Data
Imports ClsWebVer
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Collections.Generic

Partial Class FlagSRRReport
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

            Dim ErrorList As New StringBuilder
            Dim dt As New DataTable
            Dim dsData As New DataSet

            dt = ObjFungsi.GetECommerceList(Me.GetType().Name, Session("UserCode"))
            If dt.TableName.ToUpper <> "ERROR" Then

                ObjFungsi.ddlBound(ddlECom, dt, "Display", "Value")

                ViewState("ECommerceList") = dt

                dt.DefaultView.RowFilter = "Value <> ''"
                dt = dt.DefaultView.ToTable
                ObjFungsi.chkBound(chkList51, dt, "Display", "Value")

            Else
                ErrorList.AppendLine("ERROR GetECommerceList!")
            End If

            Dim ErrorMessage As String = ""
            dt = ObjFungsi.GetAccountCategory(ErrorMessage)
            If dt.TableName.ToUpper <> "ERROR" Then
                ObjFungsi.ddlBound(ddlTipeSeller, dt, "Display", "Value")

                ViewState("AccountCategory") = dt

                dt.DefaultView.RowFilter = "Value <> ''"
                dt = dt.DefaultView.ToTable
                ObjFungsi.chkBound(chkList61, dt, "Display", "Value")

            Else
                ErrorList.AppendLine("ERROR GetAccountCategory!")
            End If

            ErrorMessage = ""
            dt = ObjFungsi.GetAccountSubCategory(ErrorMessage)
            If dt.TableName.ToUpper <> "ERROR" Then
                ObjFungsi.ddlBound(ddlSubKategoriSeller, dt, "Display", "Value")

                ViewState("AccountSubCategory") = dt

                dt.DefaultView.RowFilter = "Value <> ''"
                dt = dt.DefaultView.ToTable
                ObjFungsi.chkBound(chkList7, dt, "Display", "Value")

            Else
                ErrorList.AppendLine("ERROR GetAccountSubCategory!")
            End If

            dt.TableName = ""
            dt = ObjFungsi.GetECommerceListWithCategoryAndSubCategory(Me.GetType().Name, Session("UserCode"), dsData)
            If dt.TableName.ToUpper <> "ERROR" Then
                ViewState("ECommerceListWithCategoryAndSubCategory") = dt
                ViewState("dsData") = dsData
            Else
                ErrorList.AppendLine("ERROR GetECommerceListWithCategoryAndSubCategory!")
            End If

            dt.TableName = ""
            dt = ObjFungsi.GetServiceTypeList(Me.GetType().Name, "")
            If dt.TableName.ToUpper <> "ERROR" Then

                ObjFungsi.ddlBound(ddlLayanan, dt, "Display", "Value")

                dt.DefaultView.RowFilter = "Value <> ''"
                dt = dt.DefaultView.ToTable
                ObjFungsi.chkBound(chkList40, dt, "Display", "Value")

            Else
                ErrorList.AppendLine("ERROR GetServiceTypeList!")
            End If

            lblError.Text = ErrorList.ToString.Replace(vbCrLf, "<br/>")

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Protected Sub SortDropDown(ByVal Type As String)

        Try

            'ambil semua data yang berkaitan dengan dropdown
            Dim dtAccCat As DataTable = ViewState("AccountCategory")
            Dim dtAccSubCat As DataTable = ViewState("AccountSubCategory")
            Dim dtEcomm As DataTable = ViewState("ECommerceList")
            Dim dt As DataTable = ViewState("ECommerceListWithCategoryAndSubCategory")

            Dim dtData As New DataTable
            Dim dsDataAll As DataSet = ViewState("dsData")

            Dim ErrorMessage As String = ""
            Select Case Type
                Case "CATEGORY" 'ketika ddlCategory berubah
                    '1. cek apabila ddlCategory ada value yang dipilih
                    '2. apabila ada value yang dipilih maka jalankan filter untuk ddlSubCategoy yang mappingCategory = ddlCategory.selectedvalue
                    '   dan jalan kan filter ke ddlEcomm untuk data account yang SubCategorynya ada di data ddlSubCategory
                    '3. Apabila tidak ada data tidak berubah maka tetap bound dengan data awal yang diambil ketika loadData()
                    If IsNothing(ViewState("GetTipeSeller")) Then
                        ViewState("GetTipeSeller") = "X"
                    End If
                    If GetTipeSeller("CODE") <> ViewState("GetTipeSeller") Then
                        ViewState("GetTipeSeller") = GetTipeSeller("CODE")

                        If GetTipeSeller("CODE") <> "" Then
                            '-----Filter ddlSubCategory-----
                            Dim filterCategory As String = GetTipeSeller("CODE").ToString
                            Dim filterCategoryArray() As String = filterCategory.Split("|")

                            ObjFungsi.SortDropDownAccountByCategorySubCategory("CATEGORY", filterCategoryArray, ddlSubKategoriSeller, chkList7, Nothing, dsDataAll, ErrorMessage)
                            If ErrorMessage <> "" Then
                                lblError.Text = ErrorMessage
                            End If

                            '-----Filter ddlEcomm-----
                            ObjFungsi.SortDropDownAccountByCategorySubCategory("CATEGORY-ACCOUNT", filterCategoryArray, ddlECom, chkList51, dtData, dsDataAll, ErrorMessage)
                            If ErrorMessage <> "" Then
                                lblError.Text = ErrorMessage
                            End If

                            ViewState("FilteredCategoryAccount") = dtData

                        Else
                            '-----bound ddlSubCategory-----
                            ObjFungsi.ddlBound(ddlSubKategoriSeller, dtAccSubCat, "Display", "Value")
                            dtAccSubCat.DefaultView.RowFilter = "Value <> ''"
                            ObjFungsi.chkBound(chkList7, dtAccSubCat, "Display", "Value")

                            '-----bound ddl Ecomm-----
                            ObjFungsi.ddlBound(ddlECom, dtEcomm, "Display", "Value")
                            ViewState("FilteredAccount") = dt
                            dtEcomm.DefaultView.RowFilter = "Value <> ''"
                            ObjFungsi.chkBound(chkList51, dtEcomm, "Display", "Value")
                        End If

                        'mereset selection
                        ViewState("Multi_SubCategorySeller") = ObjFungsi.SetDataMulti(btnMulti7, chkList7)
                        ViewState("Multi_SubCategorySellerName") = ObjFungsi.SetDataMultiName(btnMulti7, chkList7)

                        ViewState("Multi_Ecomm") = ObjFungsi.SetDataMulti(btnMulti51, chkList51)
                        ViewState("Multi_EcommName") = ObjFungsi.SetDataMultiName(btnMulti51, chkList51)
                    End If

                Case "SUBCATEGORY" 'ketika ddlsubcategory berubah

                    'ambil data Account keseluruhan apabila data awal belum terfilter karena tidak merubah ddlCategory
                    'apabila ddlCategory berubah maka ambil data Account yang sudah terfilter ddlCategory dari viewstate

                    If IsNothing(ViewState("GetSubCategorySeller")) Then
                        ViewState("GetSubCategorySeller") = "X"
                    End If
                    If GetSubCategorySeller("CODE") <> ViewState("GetSubCategorySeller") Then
                        ViewState("GetSubCategorySeller") = GetSubCategorySeller("CODE")

                        Dim dtFilteredAccount As DataTable = dt
                        If Not IsNothing(ViewState("FilteredCategoryAccount")) Then
                            dtFilteredAccount = ViewState("FilteredCategoryAccount")
                        End If

                        'cek apabila ddlsubcategory ada value yang dipilih
                        'bila ada maka jalankan filter untuk data Account
                        'bila tidak ada tetap jalankan bound menggunakan data account yang didapat diatas
                        If GetSubCategorySeller("CODE") <> "" Then
                            Dim filterSubCategory As String = GetSubCategorySeller("CODE").ToString
                            Dim filterSubCategoryArray() As String = filterSubCategory.Split("|")

                            ObjFungsi.SortDropDownAccountByCategorySubCategory("SUBCATEGORY-ACCOUNT", filterSubCategoryArray, ddlECom, chkList51, Nothing, dsDataAll, ErrorMessage)
                            If ErrorMessage <> "" Then
                                lblError.Text = ErrorMessage
                            End If

                            ViewState("FilteredSubCategoryAccount") = dtFilteredAccount

                        Else
                            ObjFungsi.ddlBound(ddlECom, dtFilteredAccount, "DisplayAccount", "ValueAccount")

                            dtFilteredAccount.DefaultView.RowFilter = "ValueAccount <> ''"
                            ObjFungsi.chkBound(chkList51, dtFilteredAccount, "DisplayAccount", "ValueAccount")
                        End If

                        ViewState("Multi_Ecomm") = ObjFungsi.SetDataMulti(btnMulti51, chkList51)
                        ViewState("Multi_EcommName") = ObjFungsi.SetDataMultiName(btnMulti51, chkList51)
                    End If

                Case Else

            End Select

        Catch ex As Exception
            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

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

    Private Function GetTipeSeller(ByVal Tipe As String) As String

        Dim hasil1 As String = ""

        Tipe = Tipe.ToUpper

        Try

            Dim hasil2() As String = ObjFungsi.isChecked_CheckBoxList(chkList61)

            If hasil2(0) = "9" Then
                lblError.Text = hasil2(1)
            Else
                If hasil2(0) = "0" Or hasil2(0) = "2" Then
                    If Tipe = "NAME" Then
                        hasil1 = ddlTipeSeller.SelectedItem.Text
                    Else
                        hasil1 = ddlTipeSeller.SelectedValue
                    End If
                Else
                    If Tipe = "NAME" Then
                        hasil1 = ViewState("Multi_TipeSellerName")
                    Else
                        hasil1 = ViewState("Multi_TipeSeller")
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

    Private Function GetSubCategorySeller(ByVal Tipe As String) As String

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
                        hasil1 = ViewState("Multi_SubCategorySellerName")
                    Else
                        hasil1 = ViewState("Multi_SubCategorySeller")
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

    Protected Sub btnMulti1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMulti40.Click, btnMulti51.Click, btnMulti61.Click, btnMulti7.Click

        lblError.Text = ""

        txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
        txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

        Dim btn As Button = sender
        'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
        Dim btnID As String = (btn.ID).Replace("btnMulti", "")

        If btnID = "40" Then
            td40.Visible = False
            tr40.Visible = True
            ObjFungsi.SetMulti(chkList40, ViewState("Multi_JenisLayanan"))
        ElseIf btnID = "51" Then
            td51.Visible = False
            tr51.Visible = True
            ObjFungsi.SetMulti(chkList51, ViewState("Multi_Ecomm"))
        ElseIf btnID = "61" Then
            td61.Visible = False
            tr61.Visible = True
            ObjFungsi.SetMulti(chkList61, ViewState("Multi_TipeSeller"))
        ElseIf btnID = "7" Then
            td7.Visible = False
            tr7.Visible = True
            ObjFungsi.SetMulti(chkList7, ViewState("Multi_SubCategorySeller"))
        Else
        End If

    End Sub

    Protected Sub btnAll1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAll40.Click, btnAll51.Click, btnAll61.Click, btnAll7.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnAll", "")

            If btnID = "40" Then
                ObjFungsi.SelectAll(chkList40)
            ElseIf btnID = "51" Then
                ObjFungsi.SelectAll(chkList51)
            ElseIf btnID = "61" Then
                ObjFungsi.SelectAll(chkList61)
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

    Protected Sub btnNon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNon40.Click, btnNon51.Click, btnNon61.Click, btnNon7.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnNon", "")

            If btnID = "40" Then
                ObjFungsi.DeSelectAll(chkList40)
            ElseIf btnID = "51" Then
                ObjFungsi.DeSelectAll(chkList51)
            ElseIf btnID = "61" Then
                ObjFungsi.DeSelectAll(chkList61)
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

    Protected Sub btnBatal1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBatal40.Click, btnBatal51.Click, btnBatal61.Click, btnBatal7.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnBatal", "")

            If btnID = "40" Then
                td40.Visible = True
                tr40.Visible = False
                ScriptManager1.SetFocus(btnMulti40)
            ElseIf btnID = "51" Then
                td51.Visible = True
                tr51.Visible = False
                ScriptManager1.SetFocus(btnMulti51)
            ElseIf btnID = "61" Then
                td61.Visible = True
                tr61.Visible = False
                ScriptManager1.SetFocus(btnMulti61)
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

    Protected Sub btnPilih1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPilih40.Click, btnPilih51.Click, btnPilih61.Click, btnPilih7.Click

        lblError.Text = ""

        Try
            Dim btn As Button = sender
            'Dim btnID As String = (btn.ID).Substring(btn.ID.Length - 1, 1)
            Dim btnID As String = (btn.ID).Replace("btnPilih", "")

            If btnID = "40" Then

                td40.Visible = True
                tr40.Visible = False

                ViewState("Multi_JenisLayanan") = ObjFungsi.SetDataMulti(btnMulti40, chkList40)
                ViewState("Multi_JenisLayananName") = ObjFungsi.SetDataMultiName(btnMulti40, chkList40)

                ScriptManager1.SetFocus(btnMulti40)

            ElseIf btnID = "51" Then

                td51.Visible = True
                tr51.Visible = False

                ViewState("Multi_Ecomm") = ObjFungsi.SetDataMulti(btnMulti51, chkList51)
                ViewState("Multi_EcommName") = ObjFungsi.SetDataMultiName(btnMulti51, chkList51)

                ScriptManager1.SetFocus(btnMulti51)

            ElseIf btnID = "61" Then

                td61.Visible = True
                tr61.Visible = False

                ViewState("Multi_TipeSeller") = ObjFungsi.SetDataMulti(btnMulti61, chkList61)
                ViewState("Multi_TipeSellerName") = ObjFungsi.SetDataMultiName(btnMulti61, chkList61)

                SortDropDown("CATEGORY")

                ScriptManager1.SetFocus(btnMulti61)

            ElseIf btnID = "7" Then

                td7.Visible = True
                tr7.Visible = False

                ViewState("Multi_SubCategorySeller") = ObjFungsi.SetDataMulti(btnMulti7, chkList7)
                ViewState("Multi_SubCategorySellerName") = ObjFungsi.SetDataMultiName(btnMulti7, chkList7)

                SortDropDown("SUBCATEGORY")

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

    Protected Sub btnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProses.Click

        ProsesPreview(False)

    End Sub

    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreview.Click

        ProsesPreview(True)

    End Sub

    Private Function ValidasiProses(ByRef NoAWB As String) As Boolean

        Try

            lblError.Text = ""

            txtTgl1.Text = Request.Form(txtTgl1.UniqueID)
            txtTgl2.Text = Request.Form(txtTgl2.UniqueID)

            Dim ask_date As Boolean = True

            NoAWB = txtNoAWB.Text.Trim.Replace(vbCrLf, "|")
            Dim AWB() As String = NoAWB.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)

            If AWB.Length > 10 Then
                lblError.Text = "Maksimal Jumlah AWB 10 !!"
                Return False
            End If

            If AWB.Length <= 1 Then

                If AWB.Length = 1 Then
                    If AWB(0) <> "" Then
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

    Private Sub ProsesPreview(ByVal PreviewOnly As Boolean)

        Dim NoAWB As String = ""

        Try
            If ValidasiProses(NoAWB) Then

                Dim TanggalAwal As String = txtTgl1.Text
                Dim TanggalAkhir As String = txtTgl2.Text
                Dim KodeEcommerce As String = GetEcommerce("CODE")
                Dim NamaEcommerce As String = GetEcommerce("NAME")
                Dim TipeSeller As String = GetTipeSeller("CODE")
                Dim TipeSellerName As String = GetTipeSeller("NAME")
                Dim SubCategorySeller As String = GetSubCategorySeller("CODE")
                Dim SubCategorySellerName As String = GetSubCategorySeller("NAME")
                Dim KodeLayanan As String = GetJenisLayanan("CODE")
                Dim NamaLayanan As String = GetJenisLayanan("NAME")

                Proses(TanggalAwal, TanggalAkhir, KodeEcommerce, NamaEcommerce, TipeSeller, TipeSellerName, SubCategorySeller, SubCategorySellerName, KodeLayanan, NamaLayanan, NoAWB, PreviewOnly)

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Private Sub Proses(ByVal TanggalAwal As String, ByVal TanggalAkhir As String, ByVal KodeEcommerce As String, ByVal NamaEcommerce As String, ByVal TipeSeller As String, ByVal TipeSellerName As String, ByVal SubCategorySeller As String, ByVal SubCategorySellerName As String, ByVal KodeLayanan As String, ByVal NamaLayanan As String, ByVal NoAWB As String, ByVal PreviewOnly As Boolean)

        Dim User As String = "" & Session("NIK")

        If User <> "" Then

            Dim SqlQuery As String = ""
            Dim MCon As MySqlConnection = Nothing
            Dim dt As DataTable = Nothing

            Try

                SqlQuery = "call `report_historysrr`(@TanggalAwal,@TanggalAkhir,@KodeEcommerce,@TipeSeller,@SubCategorySeller,@KodeLayanan,@NoAWB,@PreviewOnly);"

                Dim SqlParam As New Dictionary(Of String, String)
                SqlParam.Add("@TanggalAwal", TanggalAwal)
                SqlParam.Add("@TanggalAkhir", TanggalAkhir)
                SqlParam.Add("@KodeEcommerce", KodeEcommerce)
                SqlParam.Add("@TipeSeller", TipeSeller)
                SqlParam.Add("@SubCategorySeller", SubCategorySeller)
                SqlParam.Add("@KodeLayanan", KodeLayanan)
                SqlParam.Add("@NoAWB", NoAWB)
                SqlParam.Add("@PreviewOnly", PreviewOnly)

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
                    HeaderTitle(2) = "KATEGORI SELLER"
                    HeaderTitle(3) = "SUBKATEGORI SELLER"
                    HeaderTitle(4) = "NAMA E-COMM"
                    HeaderTitle(5) = "JENIS LAYANAN"
                    HeaderTitle(6) = "AWB IPP"

                    HeaderContent(0) = TanggalAwal
                    HeaderContent(1) = TanggalAkhir
                    HeaderContent(2) = TipeSellerName
                    HeaderContent(3) = SubCategorySellerName
                    HeaderContent(4) = NamaEcommerce
                    HeaderContent(5) = NamaLayanan
                    HeaderContent(6) = NoAWB.Replace("|", "/")

                    Session("TitleReport") = "LAPORAN PENGEMBALIAN ATAS PERMINTAAN PENGIRIM"
                    Session("HeaderTitleReport") = HeaderTitle
                    Session("HeaderContentReport") = HeaderContent
                    Session("BodyReport") = dt

                    'Response.Write("<script language=javascript>child=window.open('Preview.aspx');</script>")
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Preview_" & AutoNum, "child=window.open('Preview.aspx');", True)

                Else

                    Dim FileName As String = "FLAGSRR_" & AutoNum
                    Dim FileExt As String = ".csv"
                    Dim FileNameExt As String = FileName & FileExt

                    Dim PesanError As String = ""

                    Dim hasil As String = ObjFungsi.DataTableToFile(dt, "|", False, FileName, FileExt, User, PesanError)

                    Dim Result() As String = hasil.Split("|")
                    If Result(0) = "1" Then

                        PesanError = ""

                        Dim FileToZip As String = "" & Result(1)
                        Dim FilesToZip As String() = FileToZip.Split("|")

                        Dim ZipFileName As String = FileName
                        Dim ZipFileExt As String = ".zip"
                        Dim ZipFileNameExt As String = ZipFileName & ZipFileExt

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

End Class
