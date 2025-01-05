
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports ReportWeb.ClsWebVerHub
Imports Microsoft.VisualBasic.ApplicationServices
Imports ReportWeb
Imports CrystalDecisions.ReportAppServer.Prompting

Partial Class AWB3PLDetil
    Inherits System.Web.UI.Page

    Dim ObjCon As New ClsConnection
    Dim ObjSql As New ClsSQL
    Dim ObjFungsi As New ClsFungsi
    Dim serv As New LocalCore
    Dim param() As Object
    Dim respon() As Object
    Dim ScriptManager1 As New ScriptManager

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ScriptManager1 = ScriptManager.GetCurrent(Me.Page)

        Dim MyPage As String = System.IO.Path.GetFileNameWithoutExtension(Request.Url.AbsolutePath).ToLower
        Session("WebCurrentPage") = MyPage

        If Not IsPostBack Then
            LoadData()
        End If

    End Sub

    Private Sub LoadData()
        Session("NIK") = "2015548000"
        Session("NameOfUser") = "Ryu"
        Session("CurrentWebCode") = "H001"

        'If Not IsNothing(Session("CurrentWebCode")) Then

        'If IsNothing(Session("SelectedRow3PL")) Then

        Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : Session Kosong, redirect ke AWB3PL.aspx"
                ObjFungsi.WriteTracelogTxt(Pesan)

        'Response.Redirect("AWB3PL.aspx", False)

        'Else

        ViewState("CurrentWebCode") = "" & Session("CurrentWebCode")

        'Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
        'Dim Pesan As String = ""
        Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Masuk LoadData"
                ObjFungsi.WriteTracelogTxt(Pesan)

                ViewState("SelectedRow3PL") = Session("SelectedRow3PL")
                Session("SelectedRow3PL") = Nothing

                ViewState("SelectedJenis3PL") = Session("SelectedJenis3PL")
                Session("SelectedJenis3PL") = Nothing

                Dim PesanError As String = ""

                Dim JenisDaftar As String = "" & ViewState("SelectedJenis3PL")

                Dim ProcessType As String = "SERAH"
                If JenisDaftar = "PUP" Then
                    ProcessType = "JEMPUT REKANAN"
                End If

                Dim dt As DataTable = ObjFungsi.GetHubExpeditionList(ViewState("CurrentWebCode"), Session("CurrentWebCode"), ProcessType, "TRUE", "1", PesanError)
                If Not IsNothing(dt) Then

                    ObjFungsi.ddlBound(ddlOtherExpedition, dt, "Display", "Value")

                    For i As Integer = 0 To dt.Columns.Count - 1

                        If dt.Columns(i).ColumnName.ToLower = "addinfo" Then

                            Dim DicAddInfoHubExpedition As New Dictionary(Of Integer, String)

                            For j As Integer = 0 To dt.Rows.Count - 1
                                DicAddInfoHubExpedition.Add(j, "" & dt.Rows(j).Item(i))
                            Next

                            ViewState("DicAddInfoHubExpedition") = DicAddInfoHubExpedition

                            Exit For

                        End If

                    Next

                    If ddlOtherExpedition.Items.Count = 2 Then
                        ddlOtherExpedition.SelectedIndex = 1
                        ddlOtherExpedition_ExecuteSelectedIndex()
                    End If

                End If

                'DropDown Asal dan Tujuan
                PesanError = ""
                dt = ObjFungsi.GetHubList2(PesanError)
                If Not IsNothing(dt) Then
                    If dt.Rows.Count > 0 Then

                        'ddlTujuan
                        ObjFungsi.ddlBound(ddlTujuan, dt, "Display", "Value")

                        'ddlAsal
                        dt.DefaultView.RowFilter = "`Value` = '" & Session("CurrentWebCode") & "'"
                        If dt.DefaultView.Count > 0 Then
                            ObjFungsi.ddlBound(ddlAsal, dt.DefaultView.ToTable, "Display", "Value")
                        Else
                            lblError.Text = "Tidak ada data Asal!"
                        End If

                        dt.DefaultView.RowFilter = Nothing

                    Else
                        lblError.Text = "Tidak ada data Asal dan Tujuan!"
                    End If
                Else
                    lblError.Text = PesanError
                End If

                dt = ViewState("SelectedRow3PL")

                ViewState("UseDestinationCode") = False
                ViewState("DestinationCode") = ""

        TrKons.Visible = True
        TrKonsCtr.Visible = True
        TrAWB.Visible = True
        TrAWBCtr.Visible = True
        TrPUP.Visible = False
        TrPUPCtr.Visible = True

        TrTglAWB3PL.Visible = False
        TrAsal.Visible = False
        TrTujuan.Visible = False
        ddlTujuan.Visible = True
        txtTujuan.Visible = True
        TrQtyVehicle.Visible = True
        TrQtyPickup.Visible = True

        TrBerat.Visible = True

        TxtTglAWB.Text = Date.Now.ToString("yyyy-MM-dd")
                CldrTglUpload.SelectedDate = Date.Now

                If JenisDaftar = "Kons" Then

                    gvData.DataSource = dt
                    gvData.DataBind()
                    TrKons.Visible = True
                    TrKonsCtr.Visible = True

                    lblCounterKons.Text = "Total : " & dt.Rows.Count

                    TrTglAWB3PL.Visible = True
                    TrAsal.Visible = True
                    TrTujuan.Visible = True
                    ddlTujuan.Visible = True
                    TrBerat.Visible = True

                ElseIf JenisDaftar = "AWB" Then

                    Dim UseDestinationCode As Boolean = dt.Columns.Contains("DestinationCode")
                    ViewState("UseDestinationCode") = UseDestinationCode

                    gvDataAWB.DataSource = dt
                    gvDataAWB.DataBind()
                    TrAWB.Visible = True
                    TrAWBCtr.Visible = True

                    lblCounterAWB.Text = "Total : " & dt.Rows.Count

                    TrTglAWB3PL.Visible = True
                    TrAsal.Visible = True
                    If UseDestinationCode Then
                        TrTujuan.Visible = False
                        txtTujuan.Visible = False
                        txtTujuan.Text = "Kode Tujuan : " & dt.Rows(0).Item("DestinationCode")
                        ViewState("DestinationCode") = "" & dt.Rows(0).Item("DestinationCode")
                    Else
                        TrTujuan.Visible = True
                        ddlTujuan.Visible = True
                    End If
                    TrBerat.Visible = True

                ElseIf JenisDaftar = "PUP" Then

                    gvDataPUP.DataSource = dt
                    gvDataPUP.DataBind()
                    TrPUP.Visible = True
                    TrPUPCtr.Visible = True

                    lblCounterPUP.Text = "Total : " & dt.Rows.Count

                    TrTglAWB3PL.Visible = True

                    TrQtyVehicle.Visible = True
                    TrQtyPickup.Visible = True

                End If

        'End If

        ' End If

    End Sub

    Private Function GetCheckedConsNum(ByRef AWB3PL As String) As String

        Dim hasil As String = ""

        Try

            Dim dt As DataTable = ViewState("SelectedRow3PL")
            Dim JenisDaftar As String = "" & ViewState("SelectedJenis3PL")

            Dim txtAWB3PL As String = txtOtherExpeditionAWB.Text
            Dim OtherExpedition As String = ddlOtherExpedition.SelectedValue
            Dim txtOtherExpeditionWeight As TextBox 'txtOtherExpeditionWeight

            Dim SB As New StringBuilder
            Dim SBAWB3PL As New StringBuilder
            SBAWB3PL.Append("SjNum,ConsNum,Destination,Weight")

            Dim index As Integer = 0

            Dim OtherExpeditionWeight As String = ""
            Dim ConsNum_AWB As String = ""

            Dim RowsCount As Integer = dt.Rows.Count - 1
            For i As Integer = 0 To RowsCount

                If index > 0 Then
                    SB.Append("|")
                End If

                If JenisDaftar = "Kons" Then
                    txtOtherExpeditionWeight = DirectCast(Me.gvData.Rows(i).FindControl("txtOtherExpeditionWeight"), TextBox)
                    OtherExpeditionWeight = txtOtherExpeditionWeight.Text
                    ConsNum_AWB = dt.Rows(i).Item("ConsNum")
                ElseIf JenisDaftar = "AWB" Then
                    txtOtherExpeditionWeight = DirectCast(Me.gvDataAWB.Rows(i).FindControl("txtOtherExpeditionWeightAWB"), TextBox)
                    OtherExpeditionWeight = txtOtherExpeditionWeight.Text
                    ConsNum_AWB = dt.Rows(i).Item("TrackNum")
                Else
                    'txtOtherExpeditionWeight = DirectCast(Me.gvDataPUP.Rows(i).FindControl("txtOtherExpeditionWeightPUP"), TextBox)
                    'OtherExpeditionWeight = "0" '2022-08-22 09:19:00, By Cucun, Jangan diisi 0 kena error Isi berat paket dengan benar
                    OtherExpeditionWeight = ""
                    ConsNum_AWB = ""
                End If

                SB.Append(dt.Rows(i).Item("ConsNum") & "," & dt.Rows(i).Item("SjNum") & "," & OtherExpedition & "," & txtAWB3PL & "," & dt.Rows(i).Item("TrackNum") & "," & OtherExpeditionWeight)

                'ConsNum,SjNum,ExpeditionCode,ExpeditionAWB,TrackNum,Weight|ConsNum,SjNum,ExpeditionCode,ExpeditionAWB,TrackNum,Weight|...

                SBAWB3PL.Append("|")
                SBAWB3PL.Append(dt.Rows(i).Item("SjNum") & "," & ConsNum_AWB & "," & dt.Rows(i).Item("DestinationName") & "," & OtherExpeditionWeight)

                index = index + 1

            Next

            hasil = SB.ToString
            AWB3PL = SBAWB3PL.ToString

            Return hasil

        Catch ex As Exception

            lblError.Text = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            Return hasil

        End Try

    End Function

    Private Function validasiSimpan(ByRef ConsNumList As String, ByRef AWB3PL As String) As Boolean

        Try

            lblError.Text = ""

            'If ddlOtherExpedition.SelectedValue = "" Then
            '    lblError.Text = lblOtherExpedition.Text & " belum dipilih!"
            '    Return False
            'End If

            'If TrTglAWB3PL.Visible Then
            '    If TxtTglAWB.Text = "" Then
            '        lblError.Text = lblTglAWB.Text & " belum dipilih!"
            '        Return False
            '    End If
            'End If

            'If TrAsal.Visible Then
            '    If ddlAsal.SelectedValue = "" Then
            '        lblError.Text = lblAsal.Text & " belum dipilih!"
            '        Return False
            '    End If
            'End If

            'If TrTujuan.Visible Then
            '    If ddlTujuan.Visible Then
            '        If ddlTujuan.SelectedValue = "" Then
            '            lblError.Text = lblTujuan.Text & " belum dipilih!"
            '            Return False
            '        End If
            '    End If
            'End If

            'If TrBerat.Visible Then

            '    If TxtBerat.Text = "" Then
            '        lblError.Text = lblBerat.Text & " tidak boleh kosong!"
            '        Return False
            '    End If

            '    If TxtBerat.Text = "0" Then
            '        lblError.Text = lblBerat.Text & " tidak boleh 0!"
            '        Return False
            '    End If

            '    '2022-08-03 14:19, By Cucun, Pa Alex minta ditambah ada ekspedisi yang berat nya boleh mengandung koma dan tidak
            '    Dim DicAddInfo As New Dictionary(Of Integer, String)

            '    DicAddInfo = ViewState("DicAddInfoHubExpedition")

            '    If Not IsNothing(DicAddInfo) Then

            '        Dim AddInfo As String = "" & DicAddInfo(ddlOtherExpedition.SelectedIndex)

            '        Dim BolehKoma As Boolean = False

            '        If AddInfo.ToUpper.Contains("RNDWGTDCM1=Y") _
            '        Or AddInfo.ToUpper.Contains("RNDWGTDCM3=Y") Then
            '            BolehKoma = True
            '        End If

            '        Dim InputanAdaKoma As Boolean = True

            '        If ObjFungsi.isNumberOnly(TxtBerat.Text) Then
            '            InputanAdaKoma = False
            '        End If

            '        If Not BolehKoma Then
            '            If InputanAdaKoma Then
            '                lblError.Text = "Ekspedisi " & ddlOtherExpedition.SelectedItem.Text & " berat harus angka bulat!"
            '                Return False
            '            End If
            '        End If

            '    End If

            'End If

            If TrQtyVehicle.Visible = True Then
                If TxtQtyVehicle.Text = "" Then
                    lblError.Text = lblQtyVehicle.Text & " tidak boleh kosong!"
                    Return False
                End If

                If TxtQtyVehicle.Text = "0" Then
                    lblError.Text = lblQtyVehicle.Text & " tidak boleh 0!"
                    Return False
                End If
            End If

            If TrQtyPickup.Visible = True Then
                If TxtQtyPickup.Text = "" Then
                    lblError.Text = lblQtyPickup.Text & " tidak boleh kosong!"
                    Return False
                End If

                If TxtQtyPickup.Text = "0" Then
                    lblError.Text = lblQtyPickup.Text & " tidak boleh 0!"
                    Return False
                End If

            End If

            'txtOtherExpeditionAWB.Text = txtOtherExpeditionAWB.Text.Trim
            'If txtOtherExpeditionAWB.ReadOnly = False Then
            '    If txtOtherExpeditionAWB.Text = "" Then
            '        lblError.Text = lblOtherExpeditionAWB.Text & " belum diisi!"
            '        Return False
            '    End If
            'End If

            ConsNumList = GetCheckedConsNum(AWB3PL)
            'If ConsNumList = "" Then
            '    lblError.Text = "Belum ada Kons yang diisi No Resi 3PL!<br/>(Kolom 3PL Kosong tidak dianggap)"
            '    Return False
            'End If

            Return True

        Catch ex As Exception

            lblError.Text = ex.Message
            Return False

        End Try

    End Function

    Protected Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSimpan.Click

        System.Threading.Thread.Sleep(DelayDisableButton)

        Dim ConsNumList As String = ""
        Dim AWB3PL As String = ""

        If validasiSimpan(ConsNumList, AWB3PL) Then

            Dim PesanError As String = ""

            Dim OtherExpedition As String = ""
            Dim OtherExpeditionName As String = ""
            Dim OtherExpeditionAWB As String = ""
            Dim HubOri As String = ""
            Dim HubDst As String = ""
            Dim OtherExpeditionWeight As String = ""
            Dim OtherExpeditionAWBDate As String = ""
            Dim AutoGenerateAWB3PL As String = ""
            Dim GeneratedAWB3PL As String = ""

            Dim JenisDaftar As String = ViewState("SelectedJenis3PL")

            Dim TipeHarga As String = ""
            Dim DestinationCode As String = ""

            Dim UseDestinationCode As Boolean = False
            If Not IsNothing(ViewState("UseDestinationCode")) Then
                If ViewState("UseDestinationCode") = True Then
                    UseDestinationCode = True
                End If
            End If

            Dim QtyVehicle As String = ""
            Dim QtyPickup As String = ""

            OtherExpedition = ddlOtherExpedition.SelectedValue
            OtherExpeditionName = ddlOtherExpedition.SelectedItem.Text
            OtherExpeditionAWB = txtOtherExpeditionAWB.Text
            HubOri = ddlAsal.SelectedValue
            'HubDst = ddlTujuan.SelectedValue
            OtherExpeditionWeight = TxtBerat.Text
            OtherExpeditionAWBDate = TxtTglAWB.Text
            AutoGenerateAWB3PL = "0"
            If CBAutoAWB3PL.Visible Then
                If CBAutoAWB3PL.Checked Then
                    AutoGenerateAWB3PL = "1"
                End If
            End If

            If JenisDaftar = "Kons" Then

                HubDst = ddlTujuan.SelectedValue
                TipeHarga = "CON"
                DestinationCode = ""

            ElseIf JenisDaftar = "AWB" Then

                HubDst = ddlTujuan.SelectedValue
                TipeHarga = "AWB"
                DestinationCode = ""

                If UseDestinationCode Then

                    HubDst = ""
                    DestinationCode = "" & ViewState("DestinationCode")

                End If

            ElseIf JenisDaftar = "" Then

                TipeHarga = "PUP"
                QtyVehicle = TxtQtyVehicle.Text
                QtyPickup = TxtQtyPickup.Text

            End If

            If SetOtherExpeditionAWB(ConsNumList, OtherExpedition, OtherExpeditionAWB, HubOri, HubDst, OtherExpeditionWeight, OtherExpeditionAWBDate, AutoGenerateAWB3PL, GeneratedAWB3PL, TipeHarga, DestinationCode, QtyVehicle, QtyPickup, PesanError) Then

                Dim Pesan As String = ""
                Dim Success As Boolean = False

                If GeneratedAWB3PL <> "" Then

                    OtherExpeditionAWB = GeneratedAWB3PL

                    'PesanError = ""

                    'If ObjFungsi.PrintAwb3plByIPP(GeneratedAWB3PL, "NORMAL", "" & Session("CurrentWebCode"), "", OtherExpedition, PesanError) Then

                    '    'CreateMessageBox("Berhasil!", "AWB3PL.aspx")

                    '    If Pesan <> "" Then
                    '        Pesan &= "\r\n"
                    '    End If

                    '    Pesan &= "Berhasil Cetak Surat Serah Terima Ekspedisi " & GeneratedAWB3PL & " !"

                    '    Success = True

                    'Else

                    '    'CreateMessageBox(PesanError)

                    '    If PesanError <> "" Then
                    '        PesanError = "\r\nPesan Error:" & PesanError
                    '    End If

                    '    If Pesan <> "" Then
                    '        Pesan &= "\r\n"
                    '    End If

                    '    Pesan &= "Gagal Cetak Surat Serah Terima Ekspedisi " & GeneratedAWB3PL & " !" & PesanError

                    'End If

                End If


                Dim SBFileToZip As New StringBuilder
                Dim FileToZip As String = ""

                Dim Confirm As String = ConfirmChange.Value

                If Confirm.ToUpper = "YES" Then

                    Dim dtAWB3PL As DataTable = ObjFungsi.ConvertStringToDatatable(AWB3PL)

                    Dim Header(11) As String
                    Header(0) = Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    Header(1) = "Bukti Proses AWB 3PL" 'DocumentName

                    Header(2) = "" 'DocumentNo
                    Header(3) = Session("NIK") 'PetugasHUBID
                    Header(4) = Session("NameOfUser") 'PetugasHUBName
                    Header(5) = ddlOtherExpedition.SelectedItem.Text '3PLName
                    Header(6) = OtherExpeditionAWB 'AWB3PL
                    Header(7) = OtherExpeditionAWBDate 'TglAWB3PL
                    Header(8) = ddlAsal.SelectedItem.Text 'Asal
                    Header(9) = ddlTujuan.SelectedItem.Text 'Tujuan
                    If UseDestinationCode Then
                        Header(9) = "-"
                    End If
                    Header(10) = OtherExpeditionWeight & " KG" 'Berat
                    Header(11) = "Berat"
                    If JenisDaftar = "PUP" Then
                        Header(10) = QtyVehicle & " x " & QtyPickup
                        Header(11) = "Jml Kend. x Pickup"
                    End If

                    PesanError = ""

                    If ObjFungsi.PrintAWB3PL(Header, dtAWB3PL, OtherExpeditionAWB, "" & ViewState("CurrentWebCode"), "" & Session("CurrentWebCode"), "" & Session("CurrentWebName"), FileToZip, PesanError) Then

                        If Pesan <> "" Then
                            Pesan &= "\r\n"
                        End If

                        Pesan &= "Berhasil Cetak Bukti Proses AWB 3PL!"

                        SBFileToZip.AppendLine(FileToZip)
                        FileToZip = ""

                        Success = True

                    Else

                        'CreateMessageBox(PesanError)

                        If PesanError <> "" Then
                            PesanError = "\r\nPesan Error:" & PesanError
                        End If

                        If Pesan <> "" Then
                            Pesan &= "\r\n"
                        End If

                        Pesan &= "Gagal Cetak Bukti Proses AWB 3PL!" & PesanError

                    End If

                Else

                    'CreateMessageBox("Berhasil!", "AWB3PL.aspx")

                End If

                If OtherExpeditionAWB <> "" Then

                    PesanError = ""

                    If ObjFungsi.PrintSerahPackingList(OtherExpeditionAWB, "NORMAL", "" & ViewState("CurrentWebCode"), "" & Session("CurrentWebCode"), "" & Session("CurrentWebName"), OtherExpedition, OtherExpeditionName, FileToZip, PesanError) Then

                        If Pesan <> "" Then
                            Pesan &= "\r\n"
                        End If

                        Pesan &= "Berhasil Cetak Packing List " & OtherExpeditionAWB & " !"

                        SBFileToZip.AppendLine(FileToZip)
                        FileToZip = ""

                        Success = True

                    Else

                        'CreateMessageBox(PesanError)

                        If PesanError <> "" Then
                            PesanError = "\r\nPesan Error:" & PesanError
                        End If

                        If Pesan <> "" Then
                            Pesan &= "\r\n"
                        End If

                        Pesan &= "Gagal Cetak Packing List " & OtherExpeditionAWB & " !" & PesanError

                    End If

                End If

                If Success Then

                    Dim DownloadGeneratedFile As String = ObjFungsi.ReadWebConfig("DownloadGeneratedFile", False)

                    If DownloadGeneratedFile = "1" Then

                        FileToZip = SBFileToZip.ToString.Replace(vbCrLf, "|")

                        Dim FilesToZip As String() = FileToZip.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)

                        Dim FileName As String = "AWB3PL_" & ViewState("CurrentWebCode") & "_" & Session("NIK") & "_" & DateTime.Now.ToString("yyMMddHHmmssfff")
                        Dim FileExt As String = ".zip"
                        Dim FileNameExt As String = FileName & FileExt

                        Dim OutputFolder As String = "Compressed"

                        Dim CreatedFilePath As String = ObjFungsi.ProsesZip(FilesToZip, OutputFolder, FileNameExt, PesanError)

                        If CreatedFilePath <> "" Then

                            Dim fileInfo As FileInfo = New FileInfo(CreatedFilePath)

                            If fileInfo.Exists Then

                                Dim DLFileName As String = Path.GetFileNameWithoutExtension(fileInfo.Name)
                                Dim DLFileExt As String = fileInfo.Extension

                                Session("DownloadPage_FileToDownload") = CreatedFilePath
                                Session("DownloadPage_DLFileName") = DLFileName
                                Session("DownloadPage_DLFileExt") = DLFileExt

                                Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

                                ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Download_" & AutoNum, "child=window.open('DownloadPage.aspx');", True)

                            End If

                        End If

                    End If

                    CreateMessageBox(Pesan, "AWB3PL.aspx")

                Else

                    If Pesan = "" Then
                        Pesan = "Tidak Cetak Bukti Proses AWB 3PL!\r\nTidak Cetak Packing List!"
                    End If

                    CreateMessageBox(Pesan)

                End If

            Else
                If PesanError <> "" Then
                    PesanError = "\r\nPesan Error:" & PesanError
                End If
                CreateMessageBox("Gagal!" & PesanError)
            End If

        End If

    End Sub

    Private Sub CreateMessageBox(ByVal Smsg As String, Optional ByVal url As String = "")

        Dim AutoNum As String = Date.Now.ToString("yyMMddHHmmssff")

        Dim urlredirect As String = ""
        If url <> "" Then
            urlredirect = "window.location.href='" & url & "';"
        End If

        Try

            ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "Alert_" & AutoNum, "alert('" & Smsg & "'); " & urlredirect, True)

        Catch ex As Exception

            lblError.Text = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

        End Try

    End Sub

#Region "webservice"

    Private Function SetOtherExpeditionAWB(ByVal ConsNumList As String, ByVal OtherExpedition As String, ByVal OtherExpeditionAWB As String, ByVal HubOri As String, ByVal HubDst As String, ByVal OtherExpeditionWeight As String, ByVal OtherExpeditionAWBDate As String, ByVal AutoGenerateAWB3PL As String, ByRef GeneratedAWB3PL As String, ByVal TipeHarga As String, ByVal DestinationCode As String, ByVal QtyVehicle As String, ByVal QtyPickup As String, ByRef PesanError As String) As Boolean

        If Not IsNothing(Session("CurrentWebCode")) Then

            'If ViewState("CurrentWebCode") <> Session("CurrentWebCode") Then
            'Session.Clear()
            'Response.Redirect("Login.aspx", False)
            'Else
            Try

                    ReDim param(15)
                    param(0) = UserWS
                    param(1) = PassWS
                    param(2) = ConsNumList

                    param(3) = OtherExpedition
                    param(4) = OtherExpeditionAWB
                    param(5) = HubOri
                    param(6) = HubDst
                    param(7) = OtherExpeditionWeight
                    param(8) = OtherExpeditionAWBDate
                    param(11) = AutoGenerateAWB3PL

                    param(12) = TipeHarga
                    'TipeHarga
                    'CON = Serah Ekspedisi
                    'AWB = Serah Antar Alamat
                    'PUE = Jemput Ekspedisi
                    'PUP = Jemput Partner

                    param(13) = DestinationCode
                    param(14) = QtyVehicle
                    param(15) = QtyPickup

                'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
                respon = serv.SetOtherExpeditionAWB(AppName, AppVersion, Session("CurrentWebCode"), param)

                    If respon(0) = "0" Then

                        If respon.Length > 3 Then
                            GeneratedAWB3PL = "" & respon(3)
                        End If

                        Return True
                    Else
                        lblError.Text = respon(1)
                        PesanError = respon(1)
                        Return False
                    End If

                    Return False

                Catch ex As Exception

                    lblError.Text = ex.Message
                    PesanError = ex.Message

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari Halaman " & Page.Title & ", Proses " & methodName & ", Error : " & ex.Message
                    ObjFungsi.WriteTracelogTxt(Pesan)

                    Return False

                End Try

            End If

        'End If

        Return False

    End Function

#End Region

    Protected Sub BtnPickDate_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If trCldr.Visible Then
            trCldr.Visible = False
            CldrTglUpload.Visible = False
        Else
            trCldr.Visible = True
            CldrTglUpload.Visible = True
        End If

    End Sub

    Protected Sub CldrTglUpload_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        TxtTglAWB.Text = CldrTglUpload.SelectedDate.ToString("yyyy-MM-dd")
        trCldr.Visible = False
        CldrTglUpload.Visible = False

    End Sub

    Private Sub ddlOtherExpedition_ExecuteSelectedIndex()

        lblError.Text = ""

        CBAutoAWB3PL.Checked = False
        CBAutoAWB3PL.Visible = False
        txtOtherExpeditionAWB.Text = ""
        txtOtherExpeditionAWB.ReadOnly = False

        If ddlOtherExpedition.SelectedValue <> "" Then

            Dim DicAddInfo As New Dictionary(Of Integer, String)

            Try

                DicAddInfo = ViewState("DicAddInfoHubExpedition")

                If Not IsNothing(DicAddInfo) Then

                    Dim AddInfo As String = "" & DicAddInfo(ddlOtherExpedition.SelectedIndex)

                    If AddInfo.ToUpper.Contains("ALLOWGENAWB3PL=Y") Then

                        CBAutoAWB3PL.Visible = True

                    End If

                End If

            Catch ex As Exception

                lblError.Text = ex.Message

            Finally
                DicAddInfo = Nothing
            End Try

        End If

    End Sub

    Protected Sub ddlOtherExpedition_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        ddlOtherExpedition_ExecuteSelectedIndex()

    End Sub

    Protected Sub CBAutoAWB3PL_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        txtOtherExpeditionAWB.ReadOnly = False
        If CBAutoAWB3PL.Checked Then
            txtOtherExpeditionAWB.ReadOnly = True
            txtOtherExpeditionAWB.Text = ""
        End If

    End Sub

End Class
