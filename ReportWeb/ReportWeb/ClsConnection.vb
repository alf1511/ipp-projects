Imports Microsoft.VisualBasic

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.Data.OleDb

Public Class ClsConnection

    Public Function SetConnExcel(ByVal Path As String) As OleDbConnection

        Dim Conn As OleDbConnection = Nothing

        Try

            Dim ConnString = String.Format(System.Configuration.ConfigurationManager.ConnectionStrings("Excel07ConString").ToString, Path)
            Return New OleDbConnection(ConnString)

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    Public Function SetConnExcelHDR(ByVal Path As String) As OleDbConnection

        Dim Conn As OleDbConnection = Nothing

        Try

            Dim ConnString = String.Format(System.Configuration.ConfigurationManager.ConnectionStrings("Excel07ConStringHDR").ToString, Path)
            Return New OleDbConnection(ConnString)

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    Public Function SetConnExcelWrite(ByVal Path As String) As OleDbConnection

        Dim Conn As OleDbConnection = Nothing

        Try

            Dim ConnString = String.Format(System.Configuration.ConfigurationManager.ConnectionStrings("Excel07ConStringWrite").ToString, Path)
            Return New OleDbConnection(ConnString)

        Catch ex As Exception

            Dim ObjFungsi As New ClsFungsi

            ObjFungsi.WriteTracelogTxt("ClsConnection SetConnExcelWrite Error, " & ex.Message)
            Return Nothing

        End Try

    End Function

    Public Function SetConnExcelWriteNoHDR(ByVal Path As String) As OleDbConnection

        Dim Conn As OleDbConnection = Nothing

        Try

            Dim ConnString = String.Format(System.Configuration.ConfigurationManager.ConnectionStrings("Excel07ConStringWriteNoHDR").ToString, Path)
            Return New OleDbConnection(ConnString)

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    Public Function SetConn_Slave1() As MySqlConnection

        Dim Conn As MySqlConnection = Nothing

        Try

            Return New MySqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("connStrSlave1").ToString)

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    Public Function SetConn_Master() As MySqlConnection

        Dim Conn As MySqlConnection = Nothing

        Try

            Return New MySqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("connStrMaster").ToString)

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    Public Function SetConn_Slave2() As MySqlConnection

        Dim Conn As MySqlConnection = Nothing

        Try

            Return New MySqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("connStrSlave2").ToString)

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

End Class
