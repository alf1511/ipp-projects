
Imports System.Data
Imports ReportWeb.ClsWebVer
Imports System.IO
Imports System.Collections.Generic
Imports MySql.Data.MySqlClient
Imports Microsoft.VisualBasic.ApplicationServices
Imports ReportWeb

Partial Class QueryTools
    Inherits System.Web.UI.Page

    Dim MyPage As String
    Dim ObjFungsi As New ClsFungsi
    Dim ObjCon As New ClsConnection
    Dim ObjSQL As New ClsSQL
    Dim ObjService As New ClsService
    Dim serv As New LocalCore
    Dim param() As Object
    Dim respon() As Object
    'Dim ScriptManager1 As New ScriptManager

    Dim PageTimeout As Integer = 3600000 'In miliseconds

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        MyPage = System.Web.VirtualPathUtility.GetFileName(Request.RawUrl).Replace(".aspx", "")
        Dim MyLi As String = "li" & MyPage

        Dim AllowedMenu As String = "" & Session("AllowedMenu")
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
        Dim dt As DataTable
        dt = ObjFungsi.ConvertStringToDatatable("Display,Value|DATABASE LIVE,dblive|DATABASE SLAVE 1,dbslave1|DATABASE SLAVE 2,dbslave2")
        ObjFungsi.ddlBound(ddlDatabase, dt, "Display", "Value")
    End Sub

    Protected Sub BtnExecute_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnExecute.Click

        TrData.Visible = False
        TrRowsCount.Visible = False
        gvData.DataSource = Nothing
        gvData.DataBind()
        lblCtr.Text = "... Rows"

        txtResult.Text = ""

        QueryTools(txtQuery.Text)

    End Sub

    Protected Sub BtnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnClear.Click
        txtQuery.Text = ""
    End Sub

    Protected Sub BtnDownload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnDownload.Click

        Download()

    End Sub

    Private Sub Download()

        lblError.Text = ""

        Try
            Dim dt As DataTable = ViewState("gvData")

            Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

            Dim DLFileName As String = "QueryTools_" & AutoNum
            Dim DLFileExt As String = ".xlsx"
            Dim DLFileNameExt As String = DLFileName & DLFileExt

            Dim PesanError As String = ""

            Dim Results As String = ObjFungsi.WriteFileExcel(dt, "QueryTools", DLFileNameExt, PesanError)
            Dim Result() As String = Results.Split("|")
            If Result(0) = "1" Then

                Dim FileToDownload As String = Result(1)

                If FileToDownload <> "" Then

                    Dim fileInfo As FileInfo = New FileInfo(FileToDownload)

                    If fileInfo.Exists Then

                        Session("DownloadPage_FileToDownload") = FileToDownload
                        Session("DownloadPage_DLFileName") = DLFileName
                        Session("DownloadPage_DLFileExt") = DLFileExt

                        ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Download_" & AutoNum, "child=window.open('DownloadPage.aspx');", True)

                    Else

                        ObjFungsi.WriteTracelogTxt("File " & FileToDownload & " not exists")
                        lblError.Text = "File " & DLFileNameExt & " not exists"

                    End If

                Else

                    ObjFungsi.WriteTracelogTxt("File " & DLFileNameExt & " fail to create!")
                    lblError.Text = "File " & DLFileNameExt & " fail to create!"

                End If

            Else
                lblError.Text = PesanError
            End If

        Catch ex As Exception
            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message
        End Try

    End Sub

    Protected Sub BtnNewTab_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnNewTab.Click

        ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "NewTab_" & Format(Date.Now, "yyMMddHHmmss"), "child=window.open('QueryTools.aspx');", True)

    End Sub

    Private Sub QueryTools(ByVal SqlQuery As String)

        Try

            ReDim param(3)
            param(0) = UserWS
            param(1) = PassWS
            param(2) = SqlQuery
            param(3) = ddlDatabase.SelectedValue

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.QueryTools(AppName, AppVersion, param)

            If respon(0) = "0" Then
                'lblError.Text = "Berhasil !!"
                txtResult.Text = respon(1)

                Dim dt As DataTable = ObjFungsi.ConvertStringToDatatableWithBarcode(respon(2))
                If Not IsNothing(dt) Then

                    Dim RowsCount As Integer = dt.Rows.Count

                    If RowsCount > 0 Then

                        gvData.DataSource = dt
                        gvData.DataBind()
                        ViewState("gvData") = dt

                        TrData.Visible = True
                        TrRowsCount.Visible = True
                        lblCtr.Text = RowsCount & " Rows"

                    End If

                End If

            Else
                'lblError.Text = respon(1)
                txtResult.Text = respon(1)
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            lblError.Text = ex.Message

        End Try

    End Sub

    Protected Sub BtnSelect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnSelect.Click
        txtQuery.Text &= "Select * "
    End Sub

    Protected Sub BtnConcat_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnConcat.Click
        txtQuery.Text &= " concat() "
    End Sub

    Protected Sub BtnGroupConcat_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnGroupConcat.Click
        txtQuery.Text &= " group_concat() "
    End Sub

    Protected Sub BtnReplace_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnReplace.Click
        txtQuery.Text &= " replace(  ,  ,  ) "
    End Sub

    Protected Sub BtnAddTime_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnAddTime.Click
        txtQuery.Text &= " addtime "
    End Sub

    Protected Sub BtnUpdTime_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnUpdTime.Click
        txtQuery.Text &= " updtime "
    End Sub

    Protected Sub BtnTrackNum_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnTrackNum.Click
        txtQuery.Text &= " tracknum "
    End Sub

    Protected Sub BtnOrderNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOrderNo.Click
        txtQuery.Text &= " orderno "
    End Sub

    Protected Sub BtnAddInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnAddInfo.Click
        txtQuery.Text &= " addinfo "
    End Sub

    Protected Sub BtnCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCode.Click
        txtQuery.Text &= " code "
    End Sub

    Protected Sub BtnName_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnName.Click
        txtQuery.Text &= " name "
    End Sub

    Protected Sub BtnLog_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnLog.Click
        txtQuery.Text &= " `log` "
    End Sub

    Protected Sub BtnFrom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFrom.Click
        txtQuery.Text &= vbCrLf & " From "
    End Sub

    Protected Sub BtnInnerJoin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnInnerJoin.Click
        txtQuery.Text &= vbCrLf & " Inner Join on "
    End Sub

    Protected Sub BtnLeftJoin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnLeftJoin.Click
        txtQuery.Text &= vbCrLf & " Left Join on "
    End Sub

    Protected Sub BtnTransaction_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnTransaction.Click
        txtQuery.Text &= vbCrLf & " `Transaction` t "
    End Sub

    Protected Sub BtnTracking_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnTracking.Click
        txtQuery.Text &= vbCrLf & " `Tracking` r "
    End Sub

    Protected Sub BtnTrcHist_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnTrcHist.Click
        txtQuery.Text &= vbCrLf & " `TrackingHistory` h "
    End Sub

    Protected Sub BtnTrcDlvrInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnTrcDlvrInfo.Click
        txtQuery.Text &= " TransactionDeliveryInfo tdi "
    End Sub

    Protected Sub BtnAutoOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnAutoOrder.Click
        txtQuery.Text &= " AutoOrderThirdPartyTransaction ao "
    End Sub

    Protected Sub BtnAutoOrdTrc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnAutoOrdTrc.Click
        txtQuery.Text &= " AutoOrderTracking aot "
    End Sub

    Protected Sub BtnAutoOrdTrcHis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnAutoOrdTrcHis.Click
        txtQuery.Text &= " AutoOrderTrackingHistory aoth "
    End Sub

    Protected Sub BtnAutoOrdCallback_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnAutoOrdCallback.Click
        txtQuery.Text &= " AutoOrderCallbackTracking aoc "
    End Sub

    Protected Sub BtnTracelog_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnTracelog.Click
        txtQuery.Text &= " tracelog tl "
    End Sub

    Protected Sub BtnReqResLog_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnReqResLog.Click
        txtQuery.Text &= " reqreslog rl "
    End Sub

    Protected Sub BtnAccount_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnAccount.Click
        txtQuery.Text &= " account a "
    End Sub

    Protected Sub BtnIdmStore_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnIdmStore.Click
        txtQuery.Text &= " indomaretstore i "
    End Sub

    Protected Sub BtnMstLogin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnMstLogin.Click
        txtQuery.Text &= " mstlogin l "
    End Sub

    Protected Sub BtnWhere_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnWhere.Click
        txtQuery.Text &= vbCrLf & " Where "
    End Sub

    Protected Sub BtnAnd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnAnd.Click
        txtQuery.Text &= " And "
    End Sub

    Protected Sub BtnOr_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOr.Click
        txtQuery.Text &= " Or "
    End Sub

    Protected Sub BtnCurdateActInAct_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCurdateActInAct.Click
        txtQuery.Text &= " (curdate() between ActiveDate and ifnull(InActiveDate,curdate())) "
    End Sub

    Protected Sub BtnEquals_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnEquals.Click
        txtQuery.Text &= " = """""
    End Sub

    Protected Sub BtnLike_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnLike.Click
        txtQuery.Text &= " Like ""%%"""
    End Sub

    Protected Sub BtnParenthesis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnParenthesis.Click
        txtQuery.Text &= " ()"
    End Sub

    Protected Sub BtnDoubleQuote_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnDoubleQuote.Click
        txtQuery.Text &= " """""
    End Sub

    Protected Sub BtnCurdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCurdate.Click
        txtQuery.Text &= " curdate()"
    End Sub

    Protected Sub BtnGroupBy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnGroupBy.Click
        txtQuery.Text &= vbCrLf & " Group By "
    End Sub

    Protected Sub BtnOrderBy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOrderBy.Click
        txtQuery.Text &= vbCrLf & " Order By "
    End Sub

    Protected Sub BtnDesc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnDesc.Click
        txtQuery.Text &= vbCrLf & " Desc "
    End Sub

    Protected Sub BtnLimit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnLimit.Click
        txtQuery.Text &= vbCrLf & " Limit "
    End Sub

    Protected Sub BtnShowTbl_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnShowTbl.Click
        txtQuery.Text = "show tables like ""%" & txtQuery.Text & "%"""
    End Sub

    Protected Sub BtnCreateTbl_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCreateTbl.Click
        txtQuery.Text = "show create table `" & txtQuery.Text & "`"
    End Sub

    Protected Sub BtnShowProc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnShowProc.Click
        txtQuery.Text = "Show procedure status where name like ""%" & txtQuery.Text & "%"""
    End Sub

    Protected Sub BtnCreateProc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCreateProc.Click
        txtQuery.Text = "Show create procedure `" & txtQuery.Text & "`"
    End Sub

    Protected Sub BtnDBDev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnDBDev.Click
        txtQuery.Text &= "iexpress_development."
    End Sub
    Protected Sub BtnShowSlave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnShowSlave.Click
        txtQuery.Text = "Show Slave Status"
    End Sub
    Protected Sub BtnTrcBalikanStaInWhoIn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnTrcBalikanStaInWhoIn.Click

        Dim TrackNumList As String = txtQuery.Text
        TrackNumList = TrackNumList.Trim.Replace(vbCrLf, "|")
        TrackNumList = TrackNumList.Trim.Replace(vbLf, "|")

        txtQuery.Text = "call TracknumBalikan_StationinWhousein ('', '" & TrackNumList & "', 0)"

    End Sub

    Protected Sub BtnConsServiceEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnConsServiceEdit.Click

        Dim ConsNumList As String = txtQuery.Text
        ConsNumList = ConsNumList.Trim.Replace(vbCrLf, "|")
        ConsNumList = ConsNumList.Trim.Replace(vbLf, "|")

        txtQuery.Text = "call sp_ConsService_Edit ('" & ConsNumList & "', '', 0)"

    End Sub

End Class
