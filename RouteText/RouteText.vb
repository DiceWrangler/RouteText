Imports Outlook = Microsoft.Office.Interop.Outlook

Module RouteText

    Dim gShutdownRequested As Boolean
    Dim gCycleInterval As Integer
    Dim gHeartbeat As Integer
    Dim gEmailToUse As String
    Dim gTestEmail As String
    Dim gAdminEmail As String
    Dim gOutput As String
    Public gRouteTextFolderName As String
    Public gNonRouteTextFolderName As String
    Public gReportItemsFolderName As String

    Public Const APP_NAME As String = "RouteText"
    Const APP_VERSION As String = "v180831"

    Public Structure EmailMessage

        Public MailItem As Outlook.MailItem
        Public FromEmailAddress As String
        Public SubjectLine As String

    End Structure

    Public Structure SentMessage
        Public MessageID As Integer
        Public BatchID As Integer
        Public TextTypeID As Integer
        Public SentByUserID As Integer
        Public SentTime As DateTime
        Public TenantID As Integer
        Public AssetID As Integer
        Public SubjectLine As String
        Public BodyText As String
    End Structure

    Public Structure Tenant
        Public TenantID As Integer
        Public FullName As String
        Public CellPhone As String
        Public PrimaryEmail As String
    End Structure

    Public Structure Asset
        Public AssetID As Integer
        Public AddressDesc As String
        Public PropertyDesc As String
    End Structure

    Public Structure SecurityUser
        Public UserID As Integer
        Public FullName As String
        Public Email As String
        Public ActiveEmployee As Char ' Y | N
    End Structure


    Const EMAIL_TO_USE_LIVE As String = "LIVE"
    Const EMAIL_TO_USE_TEST As String = "TEST"

    Const OUTPUT_SEND As String = "SEND"
    Const OUTPUT_DRAFT As String = "DRAFT"

    Const APP_STATUS_START As String = "Starting"
    Const APP_STATUS_HEARTBEAT As String = "Running"
    Const APP_STATUS_SHUTDOWN As String = "Stopping"


    Sub Main()

        Dim lError As Integer = 0

        lError = RouteTextStartup()
        If lError <> 0 Then GoTo MAIN_EXIT

        LoadAppConfigs() ' Subroutine just uses defaults if it fails for whatever reason
        InitFolders()

        ShutdownRequestClear(APP_NAME)
        gShutdownRequested = False

        Do

            MainLoop()
            Threading.Thread.Sleep(gCycleInterval)

            If Not gShutdownRequested Then gShutdownRequested = ShutdownRequestTest(APP_NAME, False)

        Loop Until gShutdownRequested

MAIN_EXIT:
        RouteTextShutdown()

    End Sub


    Sub MainLoop()

        Dim lMessage As EmailMessage
        Dim lLastMessage As SentMessage
        Dim lTenant As Tenant
        Dim lAsset As Asset
        Dim lSecurityUser As SecurityUser
        Dim lContext As String

        lMessage = GetMessage()

        If lMessage.FromEmailAddress > "" Then

            lLastMessage = GetLastSentMessageByEmail(lMessage.FromEmailAddress)

            If lLastMessage.MessageID > 0 Then

                LogMessage("Reply From: " & lMessage.FromEmailAddress)

                lTenant = GetTenant(lLastMessage.TenantID)
                lAsset = GetAsset(lLastMessage.AssetID)
                lSecurityUser = GetSecurityUser(lLastMessage.SentByUserID)

                ' add context based on batch, tenant and asset
                lContext = GetContext(lLastMessage, lTenant, lAsset, lSecurityUser)

                ' forward the message to the sender's email
                ForwardReply(lMessage, lContext, lSecurityUser)

                ' move original message to processed folder
                FileMailItem(lMessage.MailItem, True)

            Else

                LogMessage("Unknown From: " & lMessage.FromEmailAddress)

                ' not a SendText message; forward to admin and file it somewhere else
                ForwardUnknown(lMessage)
                FileMailItem(lMessage.MailItem, False)

            End If

            OutlookFlush()

        Else

            LogHeartbeat()

        End If


    End Sub


    Function GetContext(pSentMessage As SentMessage, pTenant As Tenant, pAsset As Asset, pSecurityUser As SecurityUser) As String

        Dim lContext As New System.Text.StringBuilder

        lContext.Append("Message Sent: " & pSentMessage.SentTime.ToString & vbCrLf)
        lContext.Append("Sent By: " & pSecurityUser.FullName & vbCrLf)
        lContext.Append("Subject Line: " & pSentMessage.SubjectLine & vbCrLf)

        lContext.Append(vbCrLf)
        lContext.Append(pSentMessage.BodyText & vbCrLf)
        lContext.Append(vbCrLf)

        lContext.Append("Tenant Name: " & pTenant.FullName & vbCrLf)
        lContext.Append("Cell Phone: " & pTenant.CellPhone & vbCrLf)
        lContext.Append("Primary Email: " & pTenant.PrimaryEmail & vbCrLf)

        lContext.Append(vbCrLf)
        lContext.Append("Property: " & pAsset.PropertyDesc & vbCrLf)
        lContext.Append("Address: " & pAsset.AddressDesc & vbCrLf)

        GetContext = lContext.ToString

    End Function


    Sub ForwardReply(pOriginalMessage As EmailMessage, pContext As String, pSecurityUser As SecurityUser)

        Dim lForwardedMessage As Outlook.MailItem

        lForwardedMessage = pOriginalMessage.MailItem.Forward()

        If (pSecurityUser.ActiveEmployee = "Y") And (pSecurityUser.Email > "") Then

            If gEmailToUse = EMAIL_TO_USE_LIVE Then
                lForwardedMessage.Recipients.Add(pSecurityUser.Email) ' use employee's email address
            Else
                lForwardedMessage.Recipients.Add(gTestEmail) ' use test email address
            End If

        Else
            lForwardedMessage.Recipients.Add(gAdminEmail) ' if employee is in-active or employee email address is undefined then use administrator's email address
        End If

        lForwardedMessage.Subject = "RE: " & pOriginalMessage.SubjectLine
        lForwardedMessage.Body = lForwardedMessage.Body & vbCrLf & vbCrLf & "===[ CONTEXT ]===" & vbCrLf & vbCrLf & pContext

        If (gOutput = OUTPUT_SEND) Then

            lForwardedMessage.Send()

        Else

            lForwardedMessage.Save()
            lForwardedMessage.Close(Outlook.OlInspectorClose.olSave)

        End If

    End Sub


    Sub ForwardUnknown(pOriginalMessage As EmailMessage)

        Dim lForwardedMessage As Outlook.MailItem

        lForwardedMessage = pOriginalMessage.MailItem.Forward()
        lForwardedMessage.Recipients.Add(gAdminEmail)
        lForwardedMessage.Subject = "RouteText: Unknown Message Type"

        If (gOutput = OUTPUT_SEND) Then
            lForwardedMessage.Send()
        Else
            lForwardedMessage.Save()
            lForwardedMessage.Close(Outlook.OlInspectorClose.olSave)
        End If

    End Sub


    Function RouteTextStartup() As Integer

        Dim lError As Integer = 0
        Dim lAnyError As Integer = 0

        LogMessage(APP_NAME & ": Starting up (" & APP_VERSION & ")")

        lError = DBOpen()
        If lError <> 0 Then lAnyError = -1 'If we cannot open the database, flag the error but keep going in case there are more errors during initialization

        lError = OutlookOpen()
        If lError <> 0 Then lAnyError = -1 'If we cannot open the mail client, flag the error but keep going in case there are more errors during initialization

        UpdateAppStatus(APP_STATUS_START)

        RouteTextStartup = lAnyError

    End Function


    Sub RouteTextShutdown()

        LogMessage(APP_NAME & ": Shutting down")
        UpdateAppStatus(APP_STATUS_SHUTDOWN)

        OutlookClose()
        DBClose()

    End Sub


    Sub LoadAppConfigs()

        Dim lCycleInterval As String

        lCycleInterval = GetAppConfig(APP_NAME, "Cycle_Interval", "10000")
        If IsNumeric(lCycleInterval) Then
            gCycleInterval = Val(lCycleInterval)
        Else
            gCycleInterval = 10000
        End If
        LogMessage("Config: Cycle_Interval=" & gCycleInterval.ToString)

        If UCase(GetAppConfig(APP_NAME, "Output", OUTPUT_DRAFT)) = OUTPUT_SEND Then
            gOutput = OUTPUT_SEND ' Send messages
        Else
            gOutput = OUTPUT_DRAFT ' Create draft only
        End If
        LogMessage("Config: Output=" & gOutput)

        If UCase(GetAppConfig(APP_NAME, "Email_to_Use", EMAIL_TO_USE_TEST)) = EMAIL_TO_USE_LIVE Then
            gEmailToUse = EMAIL_TO_USE_LIVE ' Send to Tenant (LIVE)
        Else
            gEmailToUse = EMAIL_TO_USE_TEST ' Send to Developer (TEST)
        End If
        LogMessage("Config: Email_to_Use=" & gEmailToUse)

        gTestEmail = GetAppConfig(APP_NAME, "Test_Email", "dicewrangler@gmail.com")  ' Default to Scott Thorne's email address
        LogMessage("Config: Test_Email=" & gTestEmail)

        gAdminEmail = GetAppConfig(APP_NAME, "Admin_Email", "dicewrangler@gmail.com")  ' Default to Scott Thorne's email address
        LogMessage("Config: Admin_Email=" & gAdminEmail)

        gRouteTextFolderName = GetAppConfig(APP_NAME, "Folder_SendText_Replies", "Junk")
        LogMessage("Config: Folder_SendText_Replies=" & gRouteTextFolderName)

        gNonRouteTextFolderName = GetAppConfig(APP_NAME, "Folder_Unrecognized", "Junk")
        LogMessage("Config: Folder_Unrecognized=" & gNonRouteTextFolderName)

        gReportItemsFolderName = GetAppConfig(APP_NAME, "Folder_Reports", "Junk")
        LogMessage("Config: Folder_Reports=" & gReportItemsFolderName)

        LogMessage(Strings.StrDup(57, "="))

    End Sub


    Sub LogMessage(pMessage As String)

        If gHeartbeat > 0 Then Console.WriteLine()
        gHeartbeat = 0

        Console.WriteLine(Now.ToLocalTime & "> " & pMessage)

    End Sub


    Sub LogHeartbeat()

        Console.Write(".")
        gHeartbeat = (gHeartbeat + 1) Mod 80
        If gHeartbeat = 0 Then Console.WriteLine() ' Line break after 80 characters

        UpdateAppStatus(APP_STATUS_HEARTBEAT)

    End Sub

End Module
