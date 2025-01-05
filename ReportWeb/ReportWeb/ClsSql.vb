Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports MySql.Data.MySqlClient
Imports ReportWeb.ClsWebVer
Imports ClsFungsi

Public Class ClsSQL

    Public Function SQLExecuteNonQuery(ByVal Mcon As MySqlConnection, ByVal SQLQuery As String, ByVal Parameters As Dictionary(Of String, String)) As Boolean

        Dim ErrorMessage As String = ""

        Return SQLExecuteNonQuery(Mcon, Nothing, SQLQuery, Parameters, ErrorMessage)

    End Function

    Public Function SQLExecuteNonQuery(ByVal Mcon As MySqlConnection, ByVal SQLQuery As String, ByVal Parameters As Dictionary(Of String, String), ByRef ErrorMessage As String) As Boolean

        Return SQLExecuteNonQuery(Mcon, Nothing, SQLQuery, Parameters, ErrorMessage)

    End Function

    Public Function SQLExecuteNonQuery(ByVal Mcon As MySqlConnection, ByVal MTrans As MySqlTransaction, ByVal SQLQuery As String, ByVal Parameters As Dictionary(Of String, String)) As Boolean

        Dim ErrorMessage As String = ""

        Return SQLExecuteNonQuery(Mcon, MTrans, SQLQuery, Parameters, ErrorMessage)

    End Function

    Public Function SQLExecuteNonQuery(ByVal Mcon As MySqlConnection, ByVal MTrans As MySqlTransaction, ByVal SQLQuery As String, ByVal Parameters As Dictionary(Of String, String), ByRef ErrorMessage As String) As Boolean

        Dim Mcom As MySqlCommand = Nothing

        Try

            Mcom = New MySqlCommand("", Mcon)

            If Mcon.State <> ConnectionState.Open Then
                Mcon.Open()
            End If

            Mcom.CommandText = SQLQuery

            If Not IsNothing(Parameters) Then

                For Each Parameter As KeyValuePair(Of String, String) In Parameters

                    Mcom.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

                Next

            End If

            If Not IsNothing(MTrans) Then
                Mcom.Transaction = MTrans
            End If

            Mcom.ExecuteNonQuery()

            Return True

        Catch ex As Exception

            Dim ObjFungsi As New ClsFungsi

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsSQL, Proses " & methodName & ", Error : " & ex.Message & ", Query : " & SQLQuery
            ObjFungsi.WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message

            Return False

        Finally
            Mcom.Dispose()
        End Try

    End Function


    Public Function SQLExecuteScalar(ByVal Mcon As MySqlConnection, ByVal SQLQuery As String, ByVal Parameters As Dictionary(Of String, String)) As Object

        Dim Mcom As MySqlCommand = Nothing

        Try

            Mcom = New MySqlCommand("", Mcon)

            If Mcon.State <> ConnectionState.Open Then
                Mcon.Open()
            End If

            Mcom.CommandText = SQLQuery

            If Not IsNothing(Parameters) Then

                For Each Parameter As KeyValuePair(Of String, String) In Parameters

                    Mcom.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

                Next

            End If

            Return Mcom.ExecuteScalar

        Catch ex As Exception

            Dim ObjFungsi As New ClsFungsi

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsSQL, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            Return 0 'dikembalikan sebagai nol, agar bisa dibaca sebagai string maupun angka

        Finally
            Mcom.Dispose()
        End Try

    End Function


    Public Function SQLInsertIntoDatatable(ByVal Mcon As MySqlConnection, ByVal SQLQuery As String, ByVal Parameters As Dictionary(Of String, String)) As DataTable

        Dim PesanError As String = ""

        Return SQLInsertIntoDatatable(Mcon, SQLQuery, Parameters, PesanError)

    End Function

    Public Function SQLInsertIntoDatatable(ByVal MCon As MySqlConnection, ByVal SQLQuery As String, ByVal Parameters As Dictionary(Of String, String), ByRef PesanError As String) As DataTable

        Dim dt As New DataTable

        Dim MCom As MySqlCommand = Nothing
        Dim MDA As MySqlDataAdapter = Nothing

        Try

            MCom = New MySqlCommand("", MCon)
            MCom.CommandTimeout = 0

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            MCom.CommandText = SQLQuery

            If Not IsNothing(Parameters) Then

                For Each Parameter As KeyValuePair(Of String, String) In Parameters

                    MCom.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

                Next

            End If

            MDA = New MySqlDataAdapter(MCom)

            MDA.Fill(dt)

            Return dt

        Catch ex As Exception

            Dim ObjFungsi As New ClsFungsi

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsSQL, Proses " & methodName & ", Error : " & ex.Message

            ObjFungsi.WriteTracelogTxt(Pesan)

            PesanError = ex.Message

            Return Nothing

        Finally
            MDA.Dispose()
            MCom.Dispose()
            dt = Nothing
        End Try

    End Function


    Public Function SQLInsertIntoDataset(ByVal Mcon As MySqlConnection, ByVal SQLQuery As String, ByVal Parameters As Dictionary(Of String, String)) As DataSet

        Dim PesanError As String = ""

        Return SQLInsertIntoDataset(Mcon, SQLQuery, Parameters, PesanError)

    End Function

    Public Function SQLInsertIntoDataset(ByVal Mcon As MySqlConnection, ByVal SQLQuery As String, ByVal Parameters As Dictionary(Of String, String), ByRef PesanError As String) As DataSet

        Dim ds As New DataSet

        Dim Mcom As MySqlCommand = Nothing
        Dim Mda As MySqlDataAdapter = Nothing

        Try

            Mcom = New MySqlCommand("", Mcon)
            Mcom.CommandTimeout = 0

            If Mcon.State <> ConnectionState.Open Then
                Mcon.Open()
            End If

            Mcom.CommandText = SQLQuery

            If Not IsNothing(Parameters) Then

                For Each Parameter As KeyValuePair(Of String, String) In Parameters

                    Mcom.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

                Next

            End If

            Mda = New MySqlDataAdapter(Mcom)

            Mda.Fill(ds)

            Return ds

        Catch ex As Exception

            Dim ObjFungsi As New ClsFungsi

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsSQL, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            PesanError = ex.Message

            Return Nothing

        Finally
            Mda.Dispose()
            Mcom.Dispose()
            ds = Nothing
        End Try

    End Function


    Public Function CekTableSQL(ByVal MCon As MySqlConnection, ByVal TbName As String) As Boolean

        Try

            Dim SqlQuery As String = "Show Tables like @TbName"

            Dim SqlParam As New Dictionary(Of String, String)
            SqlParam.Add("@TbName", TbName)

            'cek apakah ada table sql yang dimaksud
            If SQLExecuteScalar(MCon, SqlQuery, SqlParam) & "" <> "" Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception

            Dim ObjFungsi As New ClsFungsi

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsSQL, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            Return False

        End Try

    End Function

    Public Function CekColumnFromTableSQL(ByVal MCon As MySqlConnection, ByVal TbName As String, ByVal ColName As String) As Boolean

        Try

            Dim SQLQuery As String = ""

            SQLQuery = "Show Columns From @TbName" 'cek apakah ada kolom pada table sql yang dimaksud
            SQLQuery &= " Where Field = @ColName"

            Dim SqlParam As New Dictionary(Of String, String)
            SqlParam.Add("@TbName", TbName)
            SqlParam.Add("@ColName", ColName)

            If SQLExecuteScalar(MCon, SQLQuery, SqlParam) & "" <> "" Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception

            Dim ObjFungsi As New ClsFungsi

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsSQL, Proses " & methodName & ", Error : " & ex.Message
            ObjFungsi.WriteTracelogTxt(Pesan)

            Return False

        End Try

    End Function


    Public Sub TraceLogSql(ByVal MCon As MySqlConnection, ByVal Method As String, ByVal Log As String, ByVal Keyword As String)

        Dim MCom As MySqlCommand = Nothing

        Try

            MCom = New MySqlCommand("", MCon)

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            MCom.CommandText = "Insert Into Tracelog (UpdTime, Type, AppName, Version, Method, User, Keyword, Log)"
            MCom.CommandText &= "Values (now(), 'DEBUG', @AppName, @AppVersion, @Method, 'webreport', @Keyword, @Log)"

            MCom.Parameters.Clear()
            MCom.Parameters.AddWithValue("@AppName", AppName)
            MCom.Parameters.AddWithValue("@AppVersion", AppVersion)
            MCom.Parameters.AddWithValue("@Method", Method)
            MCom.Parameters.AddWithValue("@Keyword", Keyword)
            MCom.Parameters.AddWithValue("@Log", Log)

            MCom.ExecuteNonQuery()

        Catch ex As Exception

            Dim ObjFungsi As New ClsFungsi

            Dim ErrMsg As String = "Dari ClsSQL, Proses " & Method & ", Error : " & ex.Message & " " & ex.StackTrace
            Try
                ErrMsg &= ", Query : " & MCom.CommandText
            Catch ex1 As Exception
            End Try

            ObjFungsi.WriteTracelogTxt(ErrMsg)

        Finally
            MCom.Dispose()

        End Try

    End Sub

    Public Sub FileDownloadLogSql(ByVal MCon As MySqlConnection, ByVal Method As String, ByVal PageName As String, ByVal User As String, ByVal FileType As String, ByVal LogDesc As String)

        Dim MCom As MySqlCommand = Nothing

        Try

            MCom = New MySqlCommand("", MCon)

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            MCom.CommandText = "Insert Into FileDownloadLog (UpdTime, `Type`, `User`, PageName, `Description`)"
            MCom.CommandText &= "Values (now(), @FileType, @User, @PageName, @LogDesc)"

            MCom.Parameters.Clear()
            MCom.Parameters.AddWithValue("@FileType", FileType)
            MCom.Parameters.AddWithValue("@User", User)
            MCom.Parameters.AddWithValue("@PageName", PageName)
            MCom.Parameters.AddWithValue("@LogDesc", LogDesc)

            MCom.ExecuteNonQuery()

        Catch ex As Exception

            Dim ObjFungsi As New ClsFungsi

            Dim ErrMsg As String = "Dari ClsSQL, Proses " & Method & ", Error : " & ex.Message & " " & ex.StackTrace
            Try
                ErrMsg &= ", Query : " & MCom.CommandText
            Catch ex1 As Exception
            End Try

            ObjFungsi.WriteTracelogTxt(ErrMsg)

        Finally
            MCom.Dispose()

        End Try

    End Sub

    Private Sub ErrorSQL(ByVal MCon As MySqlConnection, ByVal Method As String, ByVal Log As String _
    , Optional ByVal Keyword As String = "", Optional ByVal ShowMsg As Boolean = False)

        On Error Resume Next

        Dim Query As String

        Query = "Insert Into Tracelog ("
        Query &= " UpdTime, `Type`, `Method`, `Keyword`, `Log`"
        Query &= " ) values ("
        Query &= " now(), 'Error SQL', @method, @keyword, @logstr"
        Query &= " )"

        If MCon.State <> Data.ConnectionState.Open Then
            MCon.Open()
        End If

        Dim MCom As New MySqlCommand(Query, MCon)

        Dim Parameters As New Dictionary(Of String, String)
        Parameters.Add("@method", Method)
        Parameters.Add("@keyword", Keyword.Replace("'", "''"))
        Parameters.Add("@logstr", Log.Replace("'", "''"))

        If Not IsNothing(Parameters) Then

            For Each Parameter As KeyValuePair(Of String, String) In Parameters

                MCom.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

            Next

        End If

        MCom.ExecuteNonQuery()

        MCom.Dispose()

    End Sub

    Public Function ExecDatatable(ByVal MCon As MySqlConnection, ByVal Query As String, Optional ByVal ShowMsg As Boolean = False) As DataTable

        Dim dtHasil As New DataTable

        Try

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            Using MDar As MySqlDataAdapter = New MySqlDataAdapter(Query, MCon)
                MDar.Fill(dtHasil)
            End Using

        Catch ex As Exception
            ErrorSQL(MCon, "Datatable", ex.Message & " " & vbCrLf & Query, , ShowMsg)

            dtHasil = Nothing

        Finally

        End Try

        Return dtHasil

    End Function



    Public Function ExecDatatableWithParam(ByVal MCon As MySqlConnection, ByVal Query As String, ByVal Parameters As Dictionary(Of String, String), Optional ByVal ShowMsg As Boolean = False) As DataTable

        Dim dtHasil As New DataTable
        Dim MDar As MySqlDataAdapter = Nothing
        Dim MCom As MySqlCommand = Nothing

        Try

            MCom = New MySqlCommand("", MCon)
            MCom.CommandTimeout = 0

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            MCom.CommandText = Query

            If Not IsNothing(Parameters) Then

                For Each Parameter As KeyValuePair(Of String, String) In Parameters

                    MCom.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

                Next

            End If

            MDar = New MySqlDataAdapter(MCom)
            MDar.Fill(dtHasil)

        Catch ex As Exception
            ErrorSQL(MCon, "Datatable", ex.Message & " " & vbCrLf & Query, , ShowMsg)

            dtHasil = Nothing

        Finally
            MDar.Dispose()
            MCom.Dispose()

        End Try

        Return dtHasil

    End Function

    Public Function ExecScalar(ByVal MCon As MySqlConnection, ByVal Query As String, Optional ByVal ShowMsg As Boolean = False) As Object

        Dim MCom As New MySqlCommand(Query, MCon)

        Dim Obj As Object = Nothing

        Try
            If MCon.State <> Data.ConnectionState.Open Then
                MCon.Open()
            End If

            Obj = MCom.ExecuteScalar

        Catch ex As Exception
            ErrorSQL(MCon, "Scalar", ex.Message & " " & vbCrLf & Query, , ShowMsg)

            Obj = "-1"

        Finally
            MCom.Dispose()

        End Try

        Return Obj

    End Function

    Public Function ExecScalarWithParam(ByVal MCon As MySqlConnection, ByVal Query As String, ByVal Parameters As Dictionary(Of String, String), Optional ByVal ShowMsg As Boolean = False) As Object

        Dim MCom As MySqlCommand = Nothing

        Dim Obj As Object = Nothing

        Try
            MCom = New MySqlCommand("", MCon)
            MCom.CommandTimeout = 0

            If MCon.State <> Data.ConnectionState.Open Then
                MCon.Open()
            End If

            MCom.CommandText = Query

            If Not IsNothing(Parameters) Then

                For Each Parameter As KeyValuePair(Of String, String) In Parameters

                    MCom.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

                Next

            End If

            Obj = MCom.ExecuteScalar

        Catch ex As Exception
            ErrorSQL(MCon, "Scalar", ex.Message & " " & vbCrLf & Query, , ShowMsg)

            Obj = "-1"

        Finally
            MCom.Dispose()

        End Try

        Return Obj

    End Function


    Public Function ExecNonQuery(ByVal MCon As MySqlConnection, ByVal Query As String, Optional ByVal ShowMsg As Boolean = False) As Boolean

        Dim MCom As New MySqlCommand(Query, MCon)

        Try

            If MCon.State <> Data.ConnectionState.Open Then
                MCon.Open()
            End If

            If MCom.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            ErrorSQL(MCon, "Query", ex.Message & " " & vbCrLf & Query, , ShowMsg)

            Return False

        Finally
            MCom.Dispose()

        End Try

    End Function

    Public Function ExecNonQueryWithParam(ByVal MCon As MySqlConnection, ByVal Query As String, ByVal Parameters As Dictionary(Of String, String), Optional ByVal ShowMsg As Boolean = False) As Boolean

        Dim MCom As MySqlCommand = Nothing

        Try
            MCom = New MySqlCommand("", MCon)
            MCom.CommandTimeout = 0

            If MCon.State <> Data.ConnectionState.Open Then
                MCon.Open()
            End If

            MCom.CommandText = Query

            If Not IsNothing(Parameters) Then

                For Each Parameter As KeyValuePair(Of String, String) In Parameters

                    MCom.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

                Next

            End If

            If MCom.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            ErrorSQL(MCon, "Query", ex.Message & " " & vbCrLf & Query, , ShowMsg)

            Return False

        Finally
            MCom.Dispose()

        End Try

    End Function

End Class
