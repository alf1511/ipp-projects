Imports ReportWeb.ClsWebVer
Imports System.Globalization

Partial Class MasterPage
    Inherits System.Web.UI.MasterPage

    Dim ObjFungsi As New ClsFungsi

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim Name As String = ObjFungsi.ReadWebConfig("AppName")
        If Name <> "" Then
            Page.Title = Page.Title & " - " & Name
            lblAppName.Text = Name
        Else
            Page.Title = Page.Title & " - WEB REPORT"
        End If

        lblAppVer.Text = "v. " & ClsWebVer.AppVersion
        Dim WebSessionID As String = ObjFungsi.ReadWebConfig("SessionID")

        If Session("SessionID") = "" Then
            Session("SessionID") = WebSessionID
        End If

        If Session("SessionID") <> WebSessionID Then

            Session.Clear()
            Response.Redirect("Login.aspx", False)

            Dim VirtualPath As String = Request.CurrentExecutionFilePath
            Dim FileName As String = System.Web.VirtualPathUtility.GetFileName(VirtualPath)
            Session("RedirectMessage") = FileName & ", System SessionID berubah, Force Logout!"

        ElseIf Session("NameOfUser") = "" Or Session("NIK") = "" Then

            Session.Clear()
            Response.Redirect("Login.aspx", False)

            Dim VirtualPath As String = Request.CurrentExecutionFilePath
            Dim FileName As String = System.Web.VirtualPathUtility.GetFileName(VirtualPath)
            Session("RedirectMessage") = FileName & ", Nama atau NIK kosong"

        Else

            lblUserName.Text = "Selamat Datang, " & Session("NameOfUser") & " - " & Session("UserCode")

            Dim Menu As String = "" & Session("AllowedMenu")

            If Menu <> "" Then

                'Menu Performance Mulai
                Dim MenuPerformanceExists As Boolean = False
                For Each li As Control In MenuPerformanceItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuPerformanceExists = True
                        End If

                    End If
                Next

                If MenuPerformanceExists Then
                    MenuPerformance.Visible = True
                    Dim MenuPerformanceItemsSwitch As String = "" & Session("MenuPerformanceItems")
                    If MenuPerformanceItemsSwitch = "ON" Then
                        MenuPerformanceItems.Visible = True
                    Else
                        MenuPerformanceItems.Visible = False
                    End If
                End If
                'Menu Performance Selesai

                'Menu CheckMonitor Mulai
                Dim MenuCheckMonitorExists As Boolean = False
                For Each li As Control In MenuCheckMonitorItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuCheckMonitorExists = True
                        End If

                    End If
                Next

                If MenuCheckMonitorExists Then
                    MenuCheckMonitor.Visible = True
                    Dim MenuCheckMonitorItemsSwitch As String = "" & Session("MenuCheckMonitorItems")
                    If MenuCheckMonitorItemsSwitch = "ON" Then
                        MenuCheckMonitorItems.Visible = True
                    Else
                        MenuCheckMonitorItems.Visible = False
                    End If
                End If
                'Menu CheckMonitor Selesai

                'Menu DataUmum Mulai
                Dim MenuDataUmumExists As Boolean = False
                For Each li As Control In MenuDataUmumItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuDataUmumExists = True
                        End If

                    End If
                Next

                If MenuDataUmumExists Then
                    MenuDataUmum.Visible = True
                    Dim MenuDataUmumItemsSwitch As String = "" & Session("MenuDataUmumItems")
                    If MenuDataUmumItemsSwitch = "ON" Then
                        MenuDataUmumItems.Visible = True
                    Else
                        MenuDataUmumItems.Visible = False
                    End If
                End If
                'Menu DataUmum Selesai

                'Menu ReportSLA Mulai
                Dim MenuReportSLAExists As Boolean = False
                For Each li As Control In MenuReportSLAItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuReportSLAExists = True
                        End If

                    End If
                Next

                If MenuReportSLAExists Then
                    MenuReportSLA.Visible = True
                    Dim MenuReportSLAItemsSwitch As String = "" & Session("MenuReportSLAItems")
                    If MenuReportSLAItemsSwitch = "ON" Then
                        MenuReportSLAItems.Visible = True
                    Else
                        MenuReportSLAItems.Visible = False
                    End If
                End If
                'Menu ReportSLA Selesai

                'Menu Opr Mulai
                Dim MenuOprExists As Boolean = False
                For Each li As Control In MenuOprItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuOprExists = True
                        End If

                    End If
                Next

                If MenuOprExists Then
                    MenuOpr.Visible = True
                    Dim MenuOprItemsSwitch As String = "" & Session("MenuOprItems")
                    If MenuOprItemsSwitch = "ON" Then
                        MenuOprItems.Visible = True
                    Else
                        MenuOprItems.Visible = False
                    End If
                End If
                'Menu Opr Selesai

                'Menu LaporanOpr Mulai
                Dim MenuLaporanOprExists As Boolean = False
                For Each li As Control In MenuLaporanOprItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuLaporanOprExists = True
                        End If

                    End If
                Next

                If MenuLaporanOprExists Then
                    MenuLaporanOpr.Visible = True
                    Dim MenuLaporanOprItemsSwitch As String = "" & Session("MenuLaporanOprItems")
                    If MenuLaporanOprItemsSwitch = "ON" Then
                        MenuLaporanOprItems.Visible = True
                    Else
                        MenuLaporanOprItems.Visible = False
                    End If
                End If
                'Menu LaporanOpr Selesai

                'Menu Sarana Mulai
                Dim MenuSaranaExists As Boolean = False
                For Each li As Control In MenuSaranaItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuSaranaExists = True
                        End If

                    End If
                Next

                If MenuSaranaExists Then
                    MenuSarana.Visible = True
                    Dim MenuSaranaItemsSwitch As String = "" & Session("MenuSaranaItems")
                    If MenuSaranaItemsSwitch = "ON" Then
                        MenuSaranaItems.Visible = True
                    Else
                        MenuSaranaItems.Visible = False
                    End If
                End If
                'Menu Sarana Selesai

                'Menu SalesMkt Mulai
                Dim MenuSalesMktExists As Boolean = False
                For Each li As Control In MenuSalesMktItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuSalesMktExists = True
                        End If

                    End If
                Next

                If MenuSalesMktExists Then
                    MenuSalesMkt.Visible = True
                    Dim MenuSalesMktItemsSwitch As String = "" & Session("MenuSalesMktItems")
                    If MenuSalesMktItemsSwitch = "ON" Then
                        MenuSalesMktItems.Visible = True
                    Else
                        MenuSalesMktItems.Visible = False
                    End If
                End If
                'Menu SalesMkt Selesai

                'Menu Laporan HP Mulai
                Dim MenuLaporanHPExists As Boolean = False
                For Each li As Control In MenuLaporanHPItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuLaporanHPExists = True
                        End If

                    End If
                Next

                If MenuLaporanHPExists Then
                    MenuLaporanHP.Visible = True
                    Dim MenuLaporanHPItemsSwitch As String = "" & Session("MenuLaporanHPItems")
                    If MenuLaporanHPItemsSwitch = "ON" Then
                        MenuLaporanHPItems.Visible = True
                    Else
                        MenuLaporanHPItems.Visible = False
                    End If
                End If
                'Menu Laporan HP Selesai

                'Menu Perform Mulai
                Dim MenuPerformExists As Boolean = False
                For Each li As Control In MenuPerformItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuPerformExists = True
                        End If

                    End If
                Next

                If MenuPerformExists Then
                    MenuPerform.Visible = True
                    Dim MenuPerformItemsSwitch As String = "" & Session("MenuPerformItems")
                    If MenuPerformItemsSwitch = "ON" Then
                        MenuPerformItems.Visible = True
                    Else
                        MenuPerformItems.Visible = False
                    End If
                End If
                'Menu Perform Selesai

                'Menu PJB Mulai
                Dim MenuPJBExists As Boolean = False
                For Each li As Control In MenuPJBItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuPJBExists = True
                        End If

                    End If
                Next

                If MenuPJBExists Then
                    MenuPJB.Visible = True
                    Dim MenuPJBItemsSwitch As String = "" & Session("MenuPJBItems")
                    If MenuPJBItemsSwitch = "ON" Then
                        MenuPJBItems.Visible = True
                    Else
                        MenuPJBItems.Visible = False
                    End If
                End If
                'Menu PJB Selesai

                'Menu Transport Mulai
                Dim MenuTransportExists As Boolean = False
                For Each li As Control In MenuTransportItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuTransportExists = True
                        End If

                    End If
                Next

                If MenuTransportExists Then
                    MenuTransport.Visible = True
                    Dim MenuTransportSwitch As String = "" & Session("MenuTransportItems")
                    If MenuTransportSwitch = "ON" Then
                        MenuTransportItems.Visible = True
                    Else
                        MenuTransportItems.Visible = False
                    End If
                End If
                'Menu Transport Selesai

                'Menu OPR Support Mulai
                Dim MenuOPRSupportExists As Boolean = False
                For Each li As Control In MenuOPRSupportItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuOPRSupportExists = True
                        End If

                    End If
                Next

                If MenuOPRSupportExists Then
                    MenuOPRSupport.Visible = True
                    Dim MenuOPRSupportSwitch As String = "" & Session("MenuOPRSupportItems")
                    If MenuOPRSupportSwitch = "ON" Then
                        MenuOPRSupportItems.Visible = True
                    Else
                        MenuOPRSupportItems.Visible = False
                    End If
                End If
                'Menu OPR Support Selesai

                'Menu Generate Mulai
                Dim MenuGenerateExists As Boolean = False
                For Each li As Control In MenuGenerateItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuGenerateExists = True
                        End If

                    End If
                Next

                If MenuGenerateExists Then
                    MenuGenerate.Visible = True
                    Dim MenuGenerateSwitch As String = "" & Session("MenuGenerateItems")
                    If MenuGenerateSwitch = "ON" Then
                        MenuGenerateItems.Visible = True
                    Else
                        MenuGenerateItems.Visible = False
                    End If
                End If
                'Menu Generate Selesai

                'Menu PettyCash Mulai
                Dim MenuPettyCashExists As Boolean = False
                For Each li As Control In MenuPettyCashItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuPettyCashExists = True
                        End If

                    End If
                Next

                If MenuPettyCashExists Then
                    MenuPettyCash.Visible = True
                    Dim MenuPettyCashSwitch As String = "" & Session("MenuPettyCashItems")
                    If MenuPettyCashSwitch = "ON" Then
                        MenuPettyCashItems.Visible = True
                    Else
                        MenuPettyCashItems.Visible = False
                    End If
                End If
                'Menu PettyCash Selesai

                'Menu UMDReimb Mulai
                Dim MenuUMDReimbExists As Boolean = False
                For Each li As Control In MenuUMDReimbItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuUMDReimbExists = True
                        End If

                    End If
                Next

                If MenuUMDReimbExists Then
                    MenuUMDReimb.Visible = True
                    Dim MenuUMDReimbSwitch As String = "" & Session("MenuUMDReimbItems")
                    If MenuUMDReimbSwitch = "ON" Then
                        MenuUMDReimbItems.Visible = True
                    Else
                        MenuUMDReimbItems.Visible = False
                    End If
                End If
                'Menu UMDReimb Selesai

                'Menu RegSeller Mulai
                Dim MenuRegSellerExists As Boolean = False
                For Each li As Control In MenuRegSellerItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuRegSellerExists = True
                        End If

                    End If
                Next

                If MenuRegSellerExists Then
                    MenuRegSeller.Visible = True
                    Dim MenuRegSellerSwitch As String = "" & Session("MenuRegSellerItems")
                    If MenuRegSellerSwitch = "ON" Then
                        MenuRegSellerItems.Visible = True
                    Else
                        MenuRegSellerItems.Visible = False
                    End If
                End If
                'Menu RegSeller Selesai

                'Menu SettingDashboard Mulai
                Dim MenuSettingDashboardExists As Boolean = False
                For Each li As Control In MenuSettingDashboardItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuSettingDashboardExists = True
                        End If

                    End If
                Next

                If MenuSettingDashboardExists Then
                    MenuSettingDashboard.Visible = True
                    Dim MenuSettingDashboardSwitch As String = "" & Session("MenuSettingDashboardItems")
                    If MenuSettingDashboardSwitch = "ON" Then
                        MenuSettingDashboardItems.Visible = True
                    Else
                        MenuSettingDashboardItems.Visible = False
                    End If
                End If
                'Menu SettingDashboard Selesai

                'Menu 3PLInstant Mulai
                Dim Menu3PLInstantExists As Boolean = False
                For Each li As Control In Menu3PLInstantItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            Menu3PLInstantExists = True
                        End If

                    End If
                Next

                If Menu3PLInstantExists Then
                    Menu3PLInstant.Visible = True
                    Dim Menu3PLInstantSwitch As String = "" & Session("Menu3PLInstantItems")
                    If Menu3PLInstantSwitch = "ON" Then
                        Menu3PLInstantItems.Visible = True
                    Else
                        Menu3PLInstantItems.Visible = False
                    End If
                End If
                'Menu 3PLInstant Selesai

                'Menu IncentiveDriver Mulai
                Dim MenuIncentiveDriverExists As Boolean = False
                For Each li As Control In MenuIncentiveDriverItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuIncentiveDriverExists = True
                        End If

                    End If
                Next

                If MenuIncentiveDriverExists Then
                    MenuIncentiveDriver.Visible = True
                    Dim MenuIncentiveDriverSwitch As String = "" & Session("MenuIncentiveDriverItems")
                    If MenuIncentiveDriverSwitch = "ON" Then
                        MenuIncentiveDriverItems.Visible = True
                    Else
                        MenuIncentiveDriverItems.Visible = False
                    End If
                End If
                'Menu IncentiveDriver Selesai

                'Menu Pricing Mulai
                Dim MenuPricingExists As Boolean = False
                For Each li As Control In MenuPricingItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuPricingExists = True
                        End If

                    End If
                Next

                If MenuPricingExists Then
                    MenuPricing.Visible = True
                    Dim MenuPricingSwitch As String = "" & Session("MenuPricingItems")
                    If MenuPricingSwitch = "ON" Then
                        MenuPricingItems.Visible = True
                    Else
                        MenuPricingItems.Visible = False
                    End If
                End If
                'Menu Pricing Selesai

                'Menu GA Mulai
                Dim MenuGAExists As Boolean = False
                For Each li As Control In MenuGAItems.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                            MenuGAExists = True
                        End If

                    End If
                Next

                If MenuGAExists Then
                    MenuGA.Visible = True
                    Dim MenuGASwitch As String = "" & Session("MenuGAItems")
                    If MenuGASwitch = "ON" Then
                        MenuGAItems.Visible = True
                    Else
                        MenuGAItems.Visible = False
                    End If
                End If
                'Menu GA Selesai

                For Each li As Control In MyList.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                        End If

                    End If
                Next

                'Tambahan Menu Biasa diatas
                For Each li As Control In MyList2.Controls
                    If TypeOf li Is HtmlGenericControl Then

                        If Menu.Contains("," & li.ID & ",") Then
                            li.Visible = True
                        End If

                    End If
                Next

                'Beri tanda menu yang aktif - Mulai
                Dim CurrentMenu As String = "" & Session("CurrentMenu")
                If CurrentMenu <> "" Then

                    Try
                        Dim MyLink As HtmlControls.HtmlAnchor = CType(Me.FindControl(CurrentMenu), HtmlControls.HtmlAnchor)
                        MyLink.Attributes.Add("Class", "MenuActive")
                    Catch
                    End Try

                End If
                'Beri tanda menu yang aktif - Selesai

            End If

        End If

    End Sub

    Protected Sub btnLogout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLogout.Click

        Session.Clear()
        Response.Redirect("Login.aspx", False)

        Dim VirtualPath As String = Request.CurrentExecutionFilePath
        Dim FileName As String = System.Web.VirtualPathUtility.GetFileName(VirtualPath)
        Session("RedirectMessage") = FileName & ", Tekan Tombol Logout"

    End Sub

    Protected Sub LBPerformance_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBPerformance.Click

        If MenuPerformanceItems.Visible Then
            MenuPerformanceItems.Visible = False
            Session("MenuPerformanceItems") = "OFF"
        Else
            MenuPerformanceItems.Visible = True
            Session("MenuPerformanceItems") = "ON"
        End If

    End Sub

    Protected Sub LBCheckMonitor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBCheckMonitor.Click

        If MenuCheckMonitorItems.Visible Then
            MenuCheckMonitorItems.Visible = False
            Session("MenuCheckMonitorItems") = "OFF"
        Else
            MenuCheckMonitorItems.Visible = True
            Session("MenuCheckMonitorItems") = "ON"
        End If

    End Sub

    Protected Sub LBDataUmum_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBDataUmum.Click

        If MenuDataUmumItems.Visible Then
            MenuDataUmumItems.Visible = False
            Session("MenuDataUmumItems") = "OFF"
        Else
            MenuDataUmumItems.Visible = True
            Session("MenuDataUmumItems") = "ON"
        End If

    End Sub

    Protected Sub LBReportSLA_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBReportSLA.Click

        If MenuReportSLAItems.Visible Then
            MenuReportSLAItems.Visible = False
            Session("MenuReportSLAItems") = "OFF"
        Else
            MenuReportSLAItems.Visible = True
            Session("MenuReportSLAItems") = "ON"
        End If

    End Sub

    Protected Sub LBOpr_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBOpr.Click

        If MenuOprItems.Visible Then
            MenuOprItems.Visible = False
            Session("MenuOprItems") = "OFF"
        Else
            MenuOprItems.Visible = True
            Session("MenuOprItems") = "ON"
        End If

    End Sub

    Protected Sub LBLaporanOpr_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBLaporanOpr.Click

        If MenuLaporanOprItems.Visible Then
            MenuLaporanOprItems.Visible = False
            Session("MenuLaporanOprItems") = "OFF"
        Else
            MenuLaporanOprItems.Visible = True
            Session("MenuLaporanOprItems") = "ON"
        End If

    End Sub

    Protected Sub LBSarana_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBSarana.Click

        If MenuSaranaItems.Visible Then
            MenuSaranaItems.Visible = False
            Session("MenuSaranaItems") = "OFF"
        Else
            MenuSaranaItems.Visible = True
            Session("MenuSaranaItems") = "ON"
        End If

    End Sub

    Protected Sub LBSalesMkt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBSalesMkt.Click

        If MenuSalesMktItems.Visible Then
            MenuSalesMktItems.Visible = False
            Session("MenuSalesMktItems") = "OFF"
        Else
            MenuSalesMktItems.Visible = True
            Session("MenuSalesMktItems") = "ON"
        End If

    End Sub

    Protected Sub LBLaporanHP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBLaporanHP.Click

        If MenuLaporanHPItems.Visible Then
            MenuLaporanHPItems.Visible = False
            Session("MenuLaporanHPItems") = "OFF"
        Else
            MenuLaporanHPItems.Visible = True
            Session("MenuLaporanHPItems") = "ON"
        End If

    End Sub

    Protected Sub LBPerform_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBPerform.Click

        If MenuPerformItems.Visible Then
            MenuPerformItems.Visible = False
            Session("MenuPerformItems") = "OFF"
        Else
            MenuPerformItems.Visible = True
            Session("MenuPerformItems") = "ON"
        End If

    End Sub

    Protected Sub LBPJB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBPJB.Click

        If MenuPJBItems.Visible Then
            MenuPJBItems.Visible = False
            Session("MenuPJBItems") = "OFF"
        Else
            MenuPJBItems.Visible = True
            Session("MenuPJBItems") = "ON"
        End If

    End Sub

    Protected Sub LBTransport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBTransport.Click

        If MenuTransportItems.Visible Then
            MenuTransportItems.Visible = False
            Session("MenuTransportItems") = "OFF"
        Else
            MenuTransportItems.Visible = True
            Session("MenuTransportItems") = "ON"
        End If

    End Sub

    Protected Sub LBOPRSupport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBOPRSupport.Click

        If MenuOPRSupportItems.Visible Then
            MenuOPRSupportItems.Visible = False
            Session("MenuOPRSupportItems") = "OFF"
        Else
            MenuOPRSupportItems.Visible = True
            Session("MenuOPRSupportItems") = "ON"
        End If

    End Sub

    Protected Sub LBGenerate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBGenerate.Click

        If MenuGenerateItems.Visible Then
            MenuGenerateItems.Visible = False
            Session("MenuGenerateItems") = "OFF"
        Else
            MenuGenerateItems.Visible = True
            Session("MenuGenerateItems") = "ON"
        End If

    End Sub

    Protected Sub LBPettyCash_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBPettyCash.Click

        If MenuPettyCashItems.Visible Then
            MenuPettyCashItems.Visible = False
            Session("MenuPettyCashItems") = "OFF"
        Else
            MenuPettyCashItems.Visible = True
            Session("MenuPettyCashItems") = "ON"
        End If

    End Sub

    Protected Sub LBUMDReimb_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If MenuUMDReimbItems.Visible Then
            MenuUMDReimbItems.Visible = False
            Session("MenuUMDReimbItems") = "OFF"
        Else
            MenuUMDReimbItems.Visible = True
            Session("MenuUMDReimbItems") = "ON"
        End If

    End Sub

    Protected Sub LBRegSeller_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBRegSeller.Click

        If MenuRegSellerItems.Visible Then
            MenuRegSellerItems.Visible = False
            Session("MenuRegSellerItems") = "OFF"
        Else
            MenuRegSellerItems.Visible = True
            Session("MenuRegSellerItems") = "ON"
        End If

    End Sub

    Protected Sub LBSettingDashboard_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBSettingDashboard.Click

        If MenuSettingDashboardItems.Visible Then
            MenuSettingDashboardItems.Visible = False
            Session("MenuSettingDashboardItems") = "OFF"
        Else
            MenuSettingDashboardItems.Visible = True
            Session("MenuSettingDashboardItems") = "ON"
        End If

    End Sub

    Protected Sub LB3PLInstant_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LB3PLInstant.Click

        If Menu3PLInstantItems.Visible Then
            Menu3PLInstantItems.Visible = False
            Session("Menu3PLInstantItems") = "OFF"
        Else
            Menu3PLInstantItems.Visible = True
            Session("Menu3PLInstantItems") = "ON"
        End If

    End Sub

    Protected Sub LBIncentiveDriver_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBIncentiveDriver.Click

        If MenuIncentiveDriverItems.Visible Then
            MenuIncentiveDriverItems.Visible = False
            Session("MenuIncentiveDriverItems") = "OFF"
        Else
            MenuIncentiveDriverItems.Visible = True
            Session("MenuIncentiveDriverItems") = "ON"
        End If

    End Sub

    Protected Sub LBPricing_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If MenuPricingItems.Visible Then
            MenuPricingItems.Visible = False
            Session("MenuPricingItems") = "OFF"
        Else
            MenuPricingItems.Visible = True
            Session("MenuPricingItems") = "ON"
        End If

    End Sub

    Protected Sub LBGA_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If MenuGAItems.Visible Then
            MenuGAItems.Visible = False
            Session("MenuGAItems") = "OFF"
        Else
            MenuGAItems.Visible = True
            Session("MenuGAItems") = "ON"
        End If

    End Sub

End Class

