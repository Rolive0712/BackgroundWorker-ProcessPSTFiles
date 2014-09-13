Imports System.Data.SqlClient
Imports System.Configuration

Public Class DataAccess
    Private sqlConn As SqlConnection
    Private strConnectionString As String

    Sub New()
        strConnectionString = ConfigurationManager.ConnectionStrings("connectionstring").ConnectionString
        sqlConn = New SqlConnection(strConnectionString)
    End Sub

    Public Sub Add(<params>)
        

    End Sub
End Class
