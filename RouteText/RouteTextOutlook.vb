Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Runtime.InteropServices


Module RouteTextOutlook

    Dim gOutlook As Outlook.Application
    Dim gOutlookNS As Outlook.NameSpace
    Dim gInbox As Outlook.MAPIFolder
    Dim gRouteTextFolder As Outlook.MAPIFolder
    Dim gNonRouteTextFolder As Outlook.MAPIFolder
    Dim gReportItemsFolder As Outlook.MAPIFolder

    Private Const ROUTETEXT_FOLDERNAME = "RouteText_Forwarded"
    Private Const NONROUTETEXT_FOLDERNAME = "RouteText_Unrecognized"
    Private Const REPORTITEMS_FOLDERNAME = "RouteText_ReportItems"


    Function OutlookOpen() As Integer

        Dim lError As Integer = 0

        Try

            gOutlook = DirectCast(Marshal.GetActiveObject("Outlook.Application.16"), Outlook.Application) ' Is there a current instance of Outlook 2016?

        Catch

            Try

                gOutlook = New Outlook.Application ' If no then instantiate one

            Catch ex As Exception

                lError = -1 ' Flag failure to open mail client
                LogMessage("*** ERROR *** OutlookOpen.gOutlook: " & ex.ToString)

            End Try

        End Try

        If lError = 0 Then

            Try

                gOutlookNS = gOutlook.GetNamespace("MAPI")
                gOutlookNS.Logon("Outlook") ' Name of MAPI profile
                gOutlookNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) ' Initialize MAPI
                gInbox = gOutlookNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)

            Catch ex As Exception

                lError = -1 ' Flag failure to initialize MAPI
                LogMessage("*** ERROR *** OutlookOpen.gOutlookNS: " & ex.ToString)

            End Try

        End If

        OutlookOpen = lError

    End Function


    Public Sub OutlookFlush()

        Dim lSync As Object

        Try

            GC.Collect() ' Force garbage collection to free-up old RPC connections
            GC.WaitForPendingFinalizers()

            lSync = gOutlookNS.SyncObjects.Item(1)
            lSync.Start

        Catch

            ' If we can't Send/Receive then don't worry about it because it will happen eventually anyway

        Finally

            lSync = Nothing

        End Try

    End Sub


    Sub OutlookClose()

        gOutlookNS.Logoff()
        gOutlook.Quit()

        gOutlookNS = Nothing
        gOutlook = Nothing

    End Sub


    Public Function GetMessage() As EmailMessage

        Dim lItems As Outlook.Items = gInbox.Items
        Dim lMailItem As Outlook.MailItem
        Dim lReportItem As Outlook.ReportItem
        Dim lEmailMessage As EmailMessage

        lEmailMessage = Nothing

        If lItems.Count > 0 Then

            If lItems.Item(1).Class = Outlook.OlObjectClass.olMail Then ' only process MailItems

                lMailItem = lItems.Item(1)
                If lMailItem.MessageClass = "IPM.Note" Then

                    With lEmailMessage

                        .MailItem = lMailItem
                        .FromEmailAddress = lMailItem.SenderEmailAddress
                        .SubjectLine = lMailItem.Subject

                    End With

                End If

            End If

            If lItems.Item(1).Class = Outlook.OlObjectClass.olReport Then ' just file ReportItems

                lReportItem = lItems.Item(1)
                lReportItem.Move(gReportItemsFolder)

            End If

        End If

        GetMessage = lEmailMessage

        lMailItem = Nothing
        lReportItem = Nothing
        lItems = Nothing

    End Function


    Public Function GetMailItem() As Outlook.MailItem

        Dim lItems As Outlook.Items = gInbox.Items
        Dim lItem As Outlook.MailItem

        If lItems.Count > 0 Then
            lItem = lItems.Item(1)
        Else
            lItem = Nothing
        End If

        GetMailItem = lItem

    End Function


    Public Sub InitFolders()

        Try
            gRouteTextFolder = gInbox.Folders(ROUTETEXT_FOLDERNAME) ' folder must be INSIDE Inbox folder
        Catch
        End Try

        Try
            gNonRouteTextFolder = gInbox.Folders(NONROUTETEXT_FOLDERNAME) ' folder must be INSIDE Inbox folder
        Catch
        End Try

        Try
            gReportItemsFolder = gInbox.Folders(REPORTITEMS_FOLDERNAME) ' folder must be INSIDE Inbox folder
        Catch
        End Try

    End Sub


    Public Sub FileMailItem(pMailItem As Outlook.MailItem, pRouteTextMessage As Boolean)

        Try
            If pRouteTextMessage Then
                pMailItem.Move(gRouteTextFolder)
            Else
                pMailItem.Move(gNonRouteTextFolder)
            End If
        Catch
        End Try

    End Sub

End Module
