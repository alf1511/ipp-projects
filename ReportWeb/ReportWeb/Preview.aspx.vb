
Imports System.Data

Partial Class Preview
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.Title = "PREVIEW REPORT "
        Dim Index As Integer = 0
        Dim IndexC As Integer = 0

        Dim row As New HtmlTableRow()
        Dim cell As New HtmlTableCell()

        '''''Create Title Here'''''
        Dim MyTitle As String = ""

        Try
            MyTitle = "" & Session("TitleReport")
            Me.Title &= "" & Session("TitleReport")
            Session("TitleReport") = ""
        Catch
        End Try

        cell.InnerHtml = "<H3>" & MyTitle & "</H3>"
        row.Cells.Add(cell)
        tableTitle.Rows.Add(row)
        '''''''''''''''''''''''''''

        '''''Create Header Here'''''
        Dim HeaderTitle() As String = Session("HeaderTitleReport")
        Session("HeaderTitleReport") = Nothing

        Dim HeaderContent() As String = Session("HeaderContentReport")
        Session("HeaderContentReport") = Nothing

        Dim HeaderRowCount As Integer = 0

        Try
            HeaderRowCount = HeaderTitle.Length
        Catch
            HeaderRowCount = 0
        End Try

        Index = 0
        While Index < HeaderRowCount
            row = New HtmlTableRow()

            cell = New HtmlTableCell()
            cell.InnerHtml = Context.Server.HtmlDecode(HeaderTitle(Index))
            row.Cells.Add(cell)

            cell = New HtmlTableCell()
            cell.InnerText = ":"
            row.Cells.Add(cell)

            cell = New HtmlTableCell()
            cell.InnerHtml = Context.Server.HtmlDecode(HeaderContent(Index))
            row.Cells.Add(cell)

            tableHeader.Rows.Add(row)

            Index = Index + 1
        End While
        '''''''''''''''''''''''''''

        '''''Create Content Here'''''
        If Not IsNothing(Session("BodyReportDs")) Then

            'Cara Baru bisa lebih dari 1 Table Content
            tableContent.Visible = False

            Dim ds As New DataSet
            ds = Session("BodyReportDs")
            Session("BodyReportDs") = Nothing

            Dim sb As New StringBuilder

            For Each dt As DataTable In ds.Tables

                PanelContent.Controls.Add(DataTableToHTMLTable(dt, False))

            Next

            Try
                ds = Nothing
            Catch ex As Exception
            End Try

        Else

            'Cara Lama hanya bisa 1 Table Content
            If Not IsNothing(Session("BodyReport")) Then

                Dim dt As New DataTable
                dt = Session("BodyReport")
                Session("BodyReport") = Nothing

                Dim ContentRowCount As Integer = 0
                Dim ContentColumnCount As Integer = 0
                Try
                    ContentRowCount = dt.Rows.Count
                    ContentColumnCount = dt.Columns.Count
                Catch
                End Try

                Dim ReplaceMe As Boolean = False
                Dim FirstColumnName As String = ""
                Index = 0
                Dim RowNumber As Integer = 0
                While Index < ContentRowCount
                    ReplaceMe = False
                    FirstColumnName = dt.Rows(Index).Item(0).ToString
                    row = New HtmlTableRow()

                    If Index > 0 Then
                        If FirstColumnName.StartsWith("YESIM") Then
                            If FirstColumnName.StartsWith("YESIMHEADER_") Then
                                row.Style.Add("background-color", "yellow")
                                RowNumber = 0
                            ElseIf FirstColumnName.StartsWith("YESIMBLANKROW_") Then
                                row.Style.Add("background-color", "black")
                                RowNumber = 0
                            Else
                                If RowNumber Mod 2 = 0 Then
                                    row.Style.Add("background-color", "#E5E5E5")
                                Else
                                    row.Style.Add("background-color", "white")
                                End If
                            End If
                            ReplaceMe = True
                        Else
                            If RowNumber Mod 2 = 0 Then
                                row.Style.Add("background-color", "#E5E5E5")
                            Else
                                row.Style.Add("background-color", "white")
                            End If
                            ReplaceMe = False
                        End If
                    Else
                        row.Style.Add("background-color", "yellow")
                        ReplaceMe = True
                    End If

                    IndexC = 0
                    While IndexC < ContentColumnCount
                        cell = New HtmlTableCell()
                        If IndexC = 0 And ReplaceMe Then
                            If FirstColumnName.StartsWith("YESIMHEADER_") Then
                                cell.InnerHtml = Context.Server.HtmlDecode(dt.Rows(Index).Item(IndexC).ToString.Replace("YESIMHEADER_", ""))
                            ElseIf FirstColumnName.StartsWith("YESIMBLANKROW_") Then
                                cell.InnerHtml = " "
                            Else
                                cell.InnerHtml = Context.Server.HtmlDecode(dt.Rows(Index).Item(IndexC).ToString)
                            End If
                        Else
                            cell.InnerHtml = Context.Server.HtmlDecode(dt.Rows(Index).Item(IndexC).ToString)
                        End If

                        row.Cells.Add(cell)

                        IndexC = IndexC + 1
                    End While

                    tableContent.Rows.Add(row)

                    Index = Index + 1
                    RowNumber = RowNumber + 1
                End While

                Try
                    dt = Nothing
                Catch ex As Exception
                End Try

            End If

        End If
        '''''''''''''''''''''''''''

    End Sub

    Private Function DataTableToHTMLTable(ByVal dt As DataTable, ByVal includeHeaders As Boolean) As Table

        Dim tbl As New Table
        tbl.GridLines = GridLines.Both
        tbl.Style.Add("border-collapse", "collapse")

        Dim tr As New TableRow
        Dim cell As New TableCell

        Dim rows As Integer = dt.Rows.Count
        Dim cols As Integer = dt.Columns.Count

        If includeHeaders Then

            Dim htr As New TableHeaderRow
            Dim hcell As New TableHeaderCell
            For i As Integer = 0 To cols - 1
                hcell = New TableHeaderCell
                hcell.Text = dt.Columns(i).ColumnName.ToString
                htr.Cells.Add(hcell)
            Next
            tbl.Rows.Add(htr)

        End If

        Dim FirstColumnData As String = ""
        For j As Integer = 0 To rows - 1

            FirstColumnData = dt.Rows(j)(0).ToString

            tr = New TableRow

            If j > 0 Then 'Warna baris

                If FirstColumnData.StartsWith("YESIMBLANKROW_") Then
                    tr.Style.Add("background-color", "black")
                Else

                    If j Mod 2 = 0 Then
                        tr.Style.Add("background-color", "#E5E5E5")
                    Else
                        tr.Style.Add("background-color", "white")
                    End If

                End If

            Else
                'Warna Header
                If includeHeaders Then
                    'Warna sama dengan baris pertama
                    tr.Style.Add("background-color", "white")
                Else
                    'Buat Warna header sendiri
                    tr.Style.Add("background-color", "yellow")
                End If
            End If

            For k As Integer = 0 To cols - 1

                cell = New TableCell

                If k > 0 Then
                    cell.Text = dt.Rows(j)(k).ToString
                Else

                    If dt.Rows(j)(k).ToString.StartsWith("YESIMBLANKROW_") Then
                        cell.Text = ""
                    Else
                        cell.Text = dt.Rows(j)(k).ToString
                    End If

                End If

                tr.Cells.Add(cell)

            Next
            tbl.Rows.Add(tr)

        Next

        Return tbl

    End Function

End Class
