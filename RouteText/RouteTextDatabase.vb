Imports System.Data.SqlClient
Imports System.Configuration


Module RouteTextDatabase

    Dim gDBConn As SqlConnection


    Public Function DBOpen() As Integer

        Dim lError As Integer = 0
        Dim lConnectionString As String

        Try

            '            gDBConn = New SqlConnection("Initial Catalog=SendText;Data Source=localhost;Integrated Security=SSPI;MultipleActiveResultSets=True;")
            lConnectionString = ConfigurationManager.ConnectionStrings("SendText").ConnectionString
            gDBConn = New SqlConnection(lConnectionString)
            gDBConn.Open()

        Catch ex As Exception

            lError = -1 ' Flag failure to open database
            LogMessage("*** ERROR *** DBOpen: " & ex.ToString)

        End Try

        DBOpen = lError

    End Function


    Public Sub DBClose()

        Try

            gDBConn.Close()
            gDBConn.Dispose()

        Catch

            ' No worries if error encountered while closing

        Finally

            gDBConn = Nothing

        End Try


    End Sub


    Public Function GetAppConfig(pAppName As String, pConfigName As String, pDefaultValue As String) As String

        Dim lCmd As New SqlCommand
        Dim lResults As String

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "GetAppConfig"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@AppName", SqlDbType.VarChar)
            lCmd.Parameters("@AppName").Value = pAppName
            lCmd.Parameters.Add("@ConfigName", SqlDbType.VarChar)
            lCmd.Parameters("@ConfigName").Value = pConfigName

            lResults = lCmd.ExecuteScalar
            If IsNothing(lResults) Then lResults = pDefaultValue 'If configuration not defined just use default value

        Catch ex As Exception

            lResults = pDefaultValue ' If we encountered an error just use default value but log it anyway
            LogMessage("*** ERROR *** GetAppConfig: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

        GetAppConfig = lResults

    End Function


    Public Function GetLastSentMessageByEmail(pEmailAddress As String) As SentMessage

        Dim lSentMessage As SentMessage
        Dim lCmd As New SqlCommand
        Dim lReader As SqlDataReader

        lSentMessage = Nothing

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "GetSentMessagesByEmail"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@EmailAddress", SqlDbType.VarChar)
            lCmd.Parameters("@EmailAddress").Value = pEmailAddress

            lReader = lCmd.ExecuteReader(CommandBehavior.SingleRow)
            lReader.Read()

            If lReader.HasRows Then

                With lSentMessage

                    .MessageID = lReader.GetInt32(0)
                    .BatchID = lReader.GetInt32(1)
                    .TextTypeID = lReader.GetInt32(2)
                    .SentByUserID = lReader.GetInt32(3)
                    .SentTime = lReader.GetDateTime(4)
                    .TenantID = lReader.GetInt32(5)
                    .AssetID = lReader.GetInt32(6)
                    .TextToEmail = lReader.GetString(7)
                    .SubjectLine = lReader.GetString(8)
                    .BodyText = lReader.GetString(9)

                End With

            Else

                lSentMessage.MessageID = 0

            End If

        Catch ex As Exception

            LogMessage("*** ERROR *** GetLastSentMessageByEmail: " & ex.ToString)

        Finally

            lReader = Nothing

            lCmd.Dispose()
            lCmd = Nothing

        End Try

        GetLastSentMessageByEmail = lSentMessage

    End Function


    Public Function GetTenant(pTenantID As Integer) As Tenant

        Dim lTenant As Tenant
        Dim lCmd As New SqlCommand
        Dim lReader As SqlDataReader

        lTenant = Nothing

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "GetTenant"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@TenantID", SqlDbType.Int)
            lCmd.Parameters("@TenantID").Value = pTenantID

            lReader = lCmd.ExecuteReader(CommandBehavior.SingleRow)
            lReader.Read()

            If lReader.HasRows Then

                With lTenant

                    .TenantID = pTenantID
                    .FullName = lReader.GetString(0)
                    .CellPhone = lReader.GetString(1)
                    .PrimaryEmail = lReader.GetString(2)

                End With

            End If

        Catch ex As Exception

            LogMessage("*** ERROR *** GetTenant: " & ex.ToString)

        Finally

            lReader = Nothing

            lCmd.Dispose()
            lCmd = Nothing

        End Try

        GetTenant = lTenant

    End Function


    Public Function GetAsset(pAssetID As Integer) As Asset

        Dim lAsset As Asset
        Dim lCmd As New SqlCommand
        Dim lReader As SqlDataReader

        lAsset = Nothing

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "GetAsset"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@AssetID", SqlDbType.Int)
            lCmd.Parameters("@AssetID").Value = pAssetID

            lReader = lCmd.ExecuteReader(CommandBehavior.SingleRow)
            lReader.Read()

            If lReader.HasRows Then

                With lAsset

                    .AssetID = pAssetID
                    .AddressDesc = lReader.GetString(0)
                    .PropertyDesc = lReader.GetString(1)

                End With

            End If

        Catch ex As Exception

            LogMessage("*** ERROR *** GetAsset: " & ex.ToString)

        Finally

            lReader = Nothing

            lCmd.Dispose()
            lCmd = Nothing

        End Try

        GetAsset = lAsset

    End Function


    Public Function GetSecurityUser(pUserID As Integer) As SecurityUser

        Dim lSecurityUser As SecurityUser
        Dim lCmd As New SqlCommand
        Dim lReader As SqlDataReader

        lSecurityUser = Nothing

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "GetSecurityUser"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@UserID", SqlDbType.Int)
            lCmd.Parameters("@UserID").Value = pUserID

            lReader = lCmd.ExecuteReader(CommandBehavior.SingleRow)
            lReader.Read()

            If lReader.HasRows Then

                With lSecurityUser

                    .UserID = pUserID
                    .FullName = lReader.GetString(0)
                    .Email = lReader.GetString(1)
                    .ActiveEmployee = lReader.GetString(2)

                End With

            End If

        Catch ex As Exception

            LogMessage("*** ERROR *** GetSecurityUser: " & ex.ToString)

        Finally

            lReader = Nothing

            lCmd.Dispose()
            lCmd = Nothing

        End Try

        GetSecurityUser = lSecurityUser

    End Function


    Public Sub ShutdownRequestClear(pAppName As String)

        Dim lCmd As New SqlCommand

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "ShutdownRequestClear"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@AppName", SqlDbType.Char)
            lCmd.Parameters("@AppName").Value = pAppName

            lCmd.ExecuteNonQuery()

        Catch ex As Exception

            LogMessage("*** ERROR *** ShutdownRequestClear: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

    End Sub


    Public Sub ShutdownRequestSet(pAppName As String)

        Dim lCmd As New SqlCommand

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "ShutdownRequestSet"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@AppName", SqlDbType.Char)
            lCmd.Parameters("@AppName").Value = pAppName

            lCmd.ExecuteNonQuery()

        Catch ex As Exception

            LogMessage("*** ERROR *** ShutdownRequestSet: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

    End Sub


    Public Function ShutdownRequestTest(pAppName As String, pDefaultValue As Boolean) As Boolean

        Dim lCmd As New SqlCommand
        Dim lResults As String

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "ShutdownRequestTest"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@AppName", SqlDbType.VarChar)
            lCmd.Parameters("@AppName").Value = pAppName

            lResults = lCmd.ExecuteScalar
            If IsNothing(lResults) Then lResults = pDefaultValue.ToString 'If configuration not defined just use default value

        Catch ex As Exception

            lResults = pDefaultValue ' If we encountered an error just use default value but log it anyway
            LogMessage("*** ERROR *** ShutdownRequestTest: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

        ShutdownRequestTest = (lResults = True.ToString)

    End Function


    Public Sub UpdateAppStatus(pAppStatus As String)

        Dim lCmd As New SqlCommand

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "UpdateAppStatus"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@AppName", SqlDbType.Char)
            lCmd.Parameters("@AppName").Value = APP_NAME
            lCmd.Parameters.Add("@AppStatus", SqlDbType.Char)
            lCmd.Parameters("@AppStatus").Value = pAppStatus

            lCmd.ExecuteNonQuery()

        Catch ex As Exception

            LogMessage("*** ERROR *** UpdateAppStatus: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

    End Sub

End Module
