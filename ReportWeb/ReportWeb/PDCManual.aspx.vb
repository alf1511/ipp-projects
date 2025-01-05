
Imports System.Data
Imports ReportWeb.ClsWebVer

Partial Class PDCManual
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

        Dim dt As DataTable = ObjFungsi.GetIndomaretDCDepoList(True)
        ObjFungsi.ddlBound(ddlDC, dt, "Display", "Value")

        ViewState("DCCode") = ""
        ViewState("StoreList") = ""
        ViewState("TrackNumList") = ""

    End Sub

    Private Function validasiCari() As Boolean

        Try

            lblError.Text = ""

            gvData.DataSource = Nothing
            gvData.DataBind()
            trData.Visible = False

            TxtNoPicking.Text = ""
            trNoPicking.Visible = False

            ViewState("DCCode") = ""
            ViewState("StoreList") = ""
            ViewState("TrackNumList") = ""

            TxtKodeToko.Text = TxtKodeToko.Text.Trim.Replace(" ", "").ToUpper
            If TxtKodeToko.Text = "" Then
                lblError.Text = lblKodeToko.Text & " masih kosong!"
                Return False
            End If

            TxtTrackNum.Text = TxtTrackNum.Text.Trim.Replace(" ", "").ToUpper
            Return True

        Catch ex As Exception

            lblError.Text = ex.Message
            Return False

        End Try


    End Function

    Protected Sub BtnCari_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCari.Click

        If validasiCari() Then

            Dim StoreList As String = TxtKodeToko.Text.Trim.Replace(vbLf, "|")
            Dim TrackNumList As String = TxtTrackNum.Text.Trim.Replace(vbLf, "|")
            Dim dt As DataTable = ObjFungsi.GetRackInStoreList(ddlDC.SelectedValue, StoreList, TrackNumList)

            If dt.TableName = "ERROR" Then
                Try
                    lblError.Text = dt.Rows(0).Item(0).ToString
                Catch
                End Try
            Else

                If dt.Rows.Count > 0 Then

                    BtnProses.Visible = True
                    gvData.DataSource = dt
                    gvData.DataBind()
                    trData.Visible = True
                    ViewState("DCCode") = ddlDC.SelectedValue
                    ViewState("StoreList") = StoreList
                    ViewState("TrackNumList") = TrackNumList

                Else

                    lblError.Text = "Tidak ada data!"

                End If

            End If

        End If

    End Sub

    Private Function validasiProses() As Boolean

        Try

            TxtNoPicking.Text = ""
            trNoPicking.Visible = False

            If ViewState("DCCode") <> ddlDC.SelectedValue Then
                lblError.Text = lblDC.Text & " berubah, Tekan tombol Cari kembali!"
                Return False
            End If

            If ViewState("StoreList") <> TxtKodeToko.Text.Trim.Replace(vbCrLf, "|") Then
                lblError.Text = lblKodeToko.Text & " berubah, Tekan tombol Cari kembali!"
                Return False
            End If

            Return True

        Catch ex As Exception

            lblError.Text = ex.Message
            Return False

        End Try

    End Function

    Protected Sub BtnProses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnProses.Click

        If validasiProses() Then

            Dim DCCode As String = "" & ViewState("DCCode")
            Dim StoreList As String = "" & ViewState("StoreList")
            Dim TrackNumList As String = "" & ViewState("TrackNumList")

            Dim dt As DataTable = ObjFungsi.DCPickingManual(DCCode, StoreList, TrackNumList)

            If dt.TableName = "ERROR" Then
                Try
                    lblError.Text = dt.Rows(0).Item(0).ToString
                Catch
                End Try
            Else

                If dt.Rows.Count > 0 Then

                    Try
                        TxtNoPicking.Text = dt.Rows(0).Item(0).ToString
                        trNoPicking.Visible = True
                        BtnProses.Visible = False
                    Catch
                    End Try

                Else

                    TxtNoPicking.Text = ""
                    trNoPicking.Visible = False
                    lblError.Text = "Tidak ada data!"

                End If

            End If

        End If

    End Sub

End Class
