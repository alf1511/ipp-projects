Imports MySql.Data.MySqlClient
Imports System.Collections.Generic

Public Class ClsError

    Private MasterMCon As MySqlConnection 'Read/Write at Primary
    Private SqlParam As New Dictionary(Of String, String)

    Public Sub New()

        Dim ConSQL As String = "" & ConfigurationManager.AppSettings("ConSQL")

        If ConSQL <> "" Then
            'webconfig
            MasterMCon = New MySqlConnection(ConSQL)
        Else
            'default
            MasterMCon = New MySqlConnection("Server=localhost;Port=3306;Database=iexpress;Uid=root;Pwd=root;")
        End If

    End Sub

    Private Sub Tracelog(ByVal MCon As MySqlConnection, ByVal Type As String _
    , ByVal AppName As String, ByVal Version As String, ByVal Method As String _
    , ByVal User As String, ByVal Log As String, Optional ByVal Keyword As String = "")

        On Error Resume Next

        Dim lCon As MySqlConnection = MasterMCon.Clone
        lCon.Open()

        If IsNothing(Keyword) Then
            Keyword = "Nothing"
        End If
        If IsDBNull(Keyword) Then
            Keyword = "DbNull"
        End If
        Keyword = ("" & Keyword)

        If IsNothing(Log) Then
            Log = "Nothing"
        End If
        If IsDBNull(Log) Then
            Log = "DbNull"
        End If
        Log = ("" & Log)

        Dim Query As String = ""
        Query = "Insert Into Tracelog ("
        Query &= " UpdTime, `Type`, AppName, Version, `Method`, `User`, `Keyword`, `Log`"
        Query &= " ) values ("
        Query &= " now(), @type, @appname, @version, @method, @user, @keyword, @logstr"
        Query &= " )"

        SqlParam = New Dictionary(Of String, String)
        SqlParam.Add("@type", Type)
        SqlParam.Add("@appname", AppName)
        SqlParam.Add("@version", Version)
        SqlParam.Add("@method", Method)
        SqlParam.Add("@user", User)
        SqlParam.Add("@keyword", Left(Keyword.Replace("'", "''").Replace("\", "\\"), 100))
        SqlParam.Add("@logstr", Left(Log.Replace("'", "''").Replace("\", "\\"), 5000))

        Dim ObjSQL As New ClsSQL
        ObjSQL.ExecNonQueryWithParam(lCon, Query, SqlParam)

        lCon.Close()
        lCon.Dispose()

    End Sub

    Public Sub DebugLog(ByVal MCon As MySqlConnection _
    , ByVal AppName As String, ByVal Version As String, ByVal Method As String _
    , ByVal User As String, ByVal Log As String, Optional ByVal Keyword As String = "")

        Tracelog(MCon, "DEBUG", AppName, Version, Method, User, Log, Keyword)

    End Sub

    Public Sub ErrorLog(ByVal MCon As MySqlConnection _
    , ByVal AppName As String, ByVal Version As String, ByVal Method As String _
    , ByVal User As String, ByVal ex As Exception, ByVal AddMsg As String, Optional ByVal Keyword As String = "")

        Tracelog(MCon, "ERROR", AppName, Version, Method, User, ex.Message & vbCrLf & ex.StackTrace & IIf(AddMsg.Trim <> "", vbCrLf & AddMsg, ""), Keyword)

    End Sub

    Public Sub RequestLog(ByVal MCon As MySqlConnection _
    , ByVal AppName As String, ByVal Version As String, ByVal Method As String _
    , ByVal User As String, ByVal Log As String, Optional ByVal Keyword As String = "")

        Tracelog(MCon, "REQ", AppName, Version, Method, User, Log, Keyword)

    End Sub

    Public Sub ResponseLog(ByVal MCon As MySqlConnection _
    , ByVal AppName As String, ByVal Version As String, ByVal Method As String _
    , ByVal User As String, ByVal Log As String, Optional ByVal Keyword As String = "")

        Tracelog(MCon, "RSP", AppName, Version, Method, User, Log, Keyword)

    End Sub

    Public Sub APIRequestLog(ByVal MCon As MySqlConnection _
        , ByVal AppName As String, ByVal Version As String, ByVal Method As String _
        , ByVal User As String, ByVal Log As String, Optional ByVal URL As String = "", Optional ByVal Keyword As String = "")

        APIReqResLog(MCon, "REQ", AppName, Version, Method, User, Log, URL, Keyword)

    End Sub

    Public Sub APIResponseLog(ByVal MCon As MySqlConnection _
    , ByVal AppName As String, ByVal Version As String, ByVal Method As String _
    , ByVal User As String, ByVal Log As String, Optional ByVal URL As String = "", Optional ByVal Keyword As String = "")

        APIReqResLog(MCon, "RSP", AppName, Version, Method, User, Log, URL, Keyword)

    End Sub

    Private Sub APIReqResLog(ByVal MCon As MySqlConnection, ByVal Type As String _
    , ByVal AppName As String, ByVal Version As String, ByVal Method As String _
    , ByVal User As String, ByVal Log As String, Optional ByVal URL As String = "", Optional ByVal Keyword As String = "")

        On Error Resume Next

        Dim lCon As MySqlConnection = MasterMCon.Clone
        lCon.Open()

        Dim Query As String = ""
        Query = "Insert Into ReqResLog ("
        Query &= " UpdTime, `Type`, AppName, Version, `Method`, `User`, `Keyword`, `Log`, `URL`"
        Query &= " ) values ("
        Query &= " now(), @type, @appname, @version, @method, @user, @keyword, @logstr, @url"
        Query &= " )"

        SqlParam = New Dictionary(Of String, String)
        SqlParam.Add("@type", Type)
        SqlParam.Add("@appname", AppName)
        SqlParam.Add("@version", Version)
        SqlParam.Add("@method", Method)
        SqlParam.Add("@user", User)
        SqlParam.Add("@keyword", Left(Keyword.Replace("'", "''").Replace("\", "\\"), 100))
        SqlParam.Add("@logstr", Log.Replace("'", "''").Replace("\", "\\"))
        SqlParam.Add("@url", URL.Replace("'", "''"))

        Dim ObjSQL As New ClsSQL
        ObjSQL.ExecNonQueryWithParam(lCon, Query, SqlParam)

        lCon.Close()
        lCon.Dispose()

    End Sub

    Public Sub CallbackLog(ByVal MCon As MySqlConnection, ByVal Type As String _
    , ByVal AppName As String, ByVal Version As String, ByVal Method As String _
    , ByVal User As String, ByVal Log As String, Optional ByVal Keyword As String = "")

        On Error Resume Next

        Dim lCon As MySqlConnection = MasterMCon.Clone
        lCon.Open()

        Dim Query As String = ""
        Query = "Insert Into CallbackLog ("
        Query &= " UpdTime, `Type`, AppName, Version, `Method`, `User`, `Keyword`, `Log`"
        Query &= " ) values ("
        Query &= " now(), @type, @appname, @version, @method, @user, @keyword, @logstr"
        Query &= " )"

        SqlParam = New Dictionary(Of String, String)
        SqlParam.Add("@type", Type)
        SqlParam.Add("@appname", AppName)
        SqlParam.Add("@version", Version)
        SqlParam.Add("@method", Method)
        SqlParam.Add("@user", User)
        SqlParam.Add("@keyword", Left(Keyword.Replace("'", "''").Replace("\", "\\"), 100))
        SqlParam.Add("@logstr", Log.Replace("'", "''").Replace("\", "\\"))

        Dim ObjSQL As New ClsSQL
        ObjSQL.ExecNonQueryWithParam(lCon, Query, SqlParam)

        lCon.Close()
        lCon.Dispose()

    End Sub

    Friend Sub ErrorLog(appName As String, appVersion As String, v As String, user As String, ex As Exception, query As String)
        Throw New NotImplementedException()
    End Sub
End Class
