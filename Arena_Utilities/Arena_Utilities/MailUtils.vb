' ---------------------------------
' --- MailUtils.vb - 01/04/2016 ---
' ---------------------------------

' ------------------------------------------------------------------------------------------
' 01/04/2016 - MJeyadarmar
'            - Added SendDBMail & Send functions to Utilities data class.
' 04/13/2012 - SBakker
'            - Added MailUtils to Utilities data class.
' ------------------------------------------------------------------------------------------

Imports Arena_DataConn
Imports System.Net.Mail
Imports System.Data.SqlClient
Imports System.Text

Public Class Mail

#Region " Constants "

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

#End Region


    Public Shared Sub Send(ByVal Recipients As String, ByVal Sender As String, ByVal Subject As String, ByVal MessageText As String)
        Dim MailType As String = "DBMail"
        Dim ReturnValue As Boolean
        Select Case MailType
            Case "DBMail"
                SendDBMail(Recipients, Sender, Subject, MessageText)
            Case "SMTPMail"
                ReturnValue = SendMail(Recipients, Sender, Subject, MessageText)
        End Select

    End Sub

    Private Shared Sub SendDBMail(ByVal Recipients As String, ByVal Sender As String, ByVal Subject As String, ByVal MessageText As String)
        Dim sb As New StringBuilder
        Dim dc As New DataConnection
        Dim cnIDRIS As SqlConnection
        Dim ReturvnValue As Integer

        If String.IsNullOrEmpty(Recipients) Then Throw New ArgumentException("Invalid Recipients - Mail not sent")

        If String.IsNullOrEmpty(MessageText) AndAlso String.IsNullOrEmpty(Subject) Then
            Throw New ArgumentException("Subject or Body of the Mail should have value - Mail not sent")
        End If

        Try
            cnIDRIS = dc.GetConnection_IDRIS
            Using cmd As SqlCommand = cnIDRIS.CreateCommand
                With sb
                    .Append("Exec msdb.[dbo].sp_send_dbmail ")
                    .Append("@recipients='" + Recipients + "'")
                    If Not String.IsNullOrEmpty(Sender) Then            ' Sender can be blank; defaults to SQLAdmin
                        .Append(", @from_address ='" + Sender + "' ")
                    End If
                    If Not String.IsNullOrEmpty(Subject) Then
                        .Append(", @subject ='" + Subject + "'")
                    End If
                    If Not String.IsNullOrEmpty(MessageText) Then
                        .Append(", @body ='" + MessageText + "'")
                    End If
                End With
                cmd.CommandText = sb.ToString
                cmd.CommandType = CommandType.Text
                ReturvnValue = cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw     '  Rethrow for now
        Finally
            dc = Nothing
            cnIDRIS = Nothing
        End Try

    End Sub

    Public Shared Function SendMail(ByVal SendTo As String, ByVal From As String, _
                                ByVal Subject As String, ByVal Body As String) As Boolean
        Dim SmtpMail As New SmtpClient()
        Dim Mail As New MailMessage()

        Try
            Mail.IsBodyHtml = False
            Mail = New MailMessage
            Mail.To.Add(SendTo)
            Mail.From = New MailAddress(From)
            Mail.Subject = Subject
            Mail.Body = Body

            'This forces the email to be sent immediately
            System.Net.ServicePointManager.MaxServicePointIdleTime = 1

            'Set Host
            SmtpMail.Host = My.Settings.MailServer

            'Send the email
            SmtpMail.Send(Mail)

            Return True

        Catch ex As Exception
            Return False

        Finally

            Mail.Dispose()
            Mail = Nothing
            SmtpMail = Nothing

        End Try

    End Function

End Class
