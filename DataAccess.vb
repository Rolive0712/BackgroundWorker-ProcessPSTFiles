Imports System.Data.SqlClient
Imports System.Configuration

Public Class DataAccess
    Private sqlConn As SqlConnection
    Private strConnectionString As String

    Sub New()
        strConnectionString = ConfigurationManager.ConnectionStrings("connectionstring").ConnectionString
        sqlConn = New SqlConnection(strConnectionString)
    End Sub

    Public Sub AddOutlookMessageInfo(ByVal TPRNbr As String, ByVal recipient As String,
                                            ByVal subject As String, ByVal filePath As String,
                                            ByVal deliveryTime As String, ByVal pstFileName As String,
                                            ByVal actulaDisplayInDB As String)
        Try
            sqlConn.Open()
            Dim recipientToSP As String
            Dim subjectToSP As String

            Dim sqlCmdObj As New SqlCommand("usp_InsertOutlookMessageInfo", sqlConn)
            sqlCmdObj.CommandType = CommandType.StoredProcedure

            Dim sqlTPRNbrParam As New SqlParameter("@TPRNbr", SqlDbType.VarChar)
            sqlTPRNbrParam.Direction = ParameterDirection.Input
            sqlTPRNbrParam.Value = TPRNbr.Trim()
            sqlCmdObj.Parameters.Add(sqlTPRNbrParam)

            Dim sqlRecipientParam As New SqlParameter("@recipients", SqlDbType.VarChar)
            sqlRecipientParam.Direction = ParameterDirection.Input
            If recipient Is Nothing Then
                recipientToSP = String.Empty
            Else
                recipientToSP = recipient.Trim()
            End If
            sqlRecipientParam.Value = recipientToSP
            sqlCmdObj.Parameters.Add(sqlRecipientParam)

            Dim sqlSubjectParam As New SqlParameter("@subject", SqlDbType.VarChar)
            sqlSubjectParam.Direction = ParameterDirection.Input
            If subject Is Nothing Then
                subjectToSP = String.Empty
            Else
                subjectToSP = subject.Trim()
            End If
            sqlSubjectParam.Value = subjectToSP
            sqlCmdObj.Parameters.Add(sqlSubjectParam)

            Dim sqlFilePathParam As New SqlParameter("@filepath", SqlDbType.VarChar)
            sqlFilePathParam.Direction = ParameterDirection.Input
            sqlFilePathParam.Value = filePath.Trim()
            sqlCmdObj.Parameters.Add(sqlFilePathParam)

            Dim sqlTimeParam As New SqlParameter("@deliveryTime", SqlDbType.VarChar)
            sqlTimeParam.Direction = ParameterDirection.Input
            sqlTimeParam.Value = deliveryTime.Trim()
            sqlCmdObj.Parameters.Add(sqlTimeParam)

            Dim sqlpstFileNameParam As New SqlParameter("@pstFileName", SqlDbType.VarChar)
            sqlpstFileNameParam.Direction = ParameterDirection.Input
            sqlpstFileNameParam.Value = pstFileName.Trim()
            sqlCmdObj.Parameters.Add(sqlpstFileNameParam)


            Dim sqldisplayNameParam As New SqlParameter("@displayName", SqlDbType.VarChar)
            sqldisplayNameParam.Direction = ParameterDirection.Input
            sqldisplayNameParam.Value = actulaDisplayInDB.Trim()
            sqlCmdObj.Parameters.Add(sqldisplayNameParam)

            sqlCmdObj.ExecuteNonQuery()
            sqlCmdObj.Parameters.Clear()

        Catch ex As Exception
            MessageBox.Show(ex.Message & "proj no:" & TPRNbr)
            'BLL.ErrorHandler.SendEmail(ex)
            '            logObj.WriteLog(ex)
            'logObj.WriteFunctionEntryLog("TPRNbr -->" + TPRNbr.ToString() + "PrjNm -->" + PrjNm.ToString() + "Prp -->" + Prp.ToString() + "PMSOEID -->" + PMSOEID.ToString() + "LOB -->" + LOB.ToString() + "RODt -->" + RODt.ToString())
        Finally
            sqlConn.Close()
        End Try


    End Sub
End Class
