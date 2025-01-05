Imports Microsoft.VisualBasic

Public Class ClsWebVerHub

    Public Const AppName As String = "WEB HUB"
    Public Const AppVersion As String = "1.7.0.5"

    Public Const UserWS As String = "webhub"
    Public Const PassWS As String = "HuB5ysT31v1"

    Public Const maxTryWS As Integer = 3

    Public Const DelayDisableButton As Integer = 150

    '0 = OK     '1 = SOBEK
    '2 = BASAH  '3 = SOBEK DAN BASAH
    Public Const DefaultKondisi As String = "0"

    '0 = OK     '1 = SOBEK
    '2 = BASAH  '3 = SOBEK DAN BASAH
    Public Const KondisiAfScn As String = "0"

    Public Const MaxAWBInCons As Integer = 100

End Class
