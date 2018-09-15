Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Runtime.InteropServices


Module RouteTextOutlook

    Dim gOutlook As Outlook.Application
    Dim gOutlookNS As Outlook.NameSpace
    Dim gInbox As Outlook.MAPIFolder
    Dim gJunk As Outlook.MAPIFolder
    Dim gRouteTextFolder As Outlook.MAPIFolder
    Dim gNonRouteTextFolder As Outlook.MAPIFolder
    Dim gReportItemsFolder As Outlook.MAPIFolder


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
                gJunk = gOutlookNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk)

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

        Catch ex As Exception

            ' If we can't Send/Receive then don't worry about it because it will happen eventually anyway
            LogMessage("*** ERROR *** OutlookFlush: " & ex.ToString)

        Finally

            lSync = Nothing

        End Try

    End Sub


    Sub OutlookClose()

        Try
            gOutlookNS.Logoff()
            gOutlook.Quit()

        Catch ex As Exception
            LogMessage("*** ERROR *** OutlookClose: " & ex.ToString)

        Finally
            gOutlookNS = Nothing
            gOutlook = Nothing

        End Try

    End Sub


    Public Function GetMessage() As EmailMessage

        Dim lItems As Outlook.Items
        Dim lItemPresent As Boolean
        Dim lItem As New Object
        Dim lMailItem As Outlook.MailItem
        Dim lReportItem As Outlook.ReportItem
        Dim lEmailMessage As EmailMessage
        'Dim lReportString As String
        'Dim lStartPosition, lEndPosition As Integer
        'Dim lOriginalRecipient As String

        lEmailMessage = Nothing

        ' *** TEST ***
        'Dim gTestFolder As Outlook.MAPIFolder
        'gTestFolder = gInbox.Folders("RouteText_TEST")
        'lItems = gTestFolder.Items
        ' ^^^ TEST ^^^

        Try
            lItems = gInbox.Items
            lItemPresent = (lItems.Count > 0)
            If lItemPresent Then lItem = lItems.Item(1) ' only process first item in folder
        Catch ex As Exception
            LogMessage("*** ERROR *** GetMessage.Items: " & ex.ToString)
            lItemPresent = False
        End Try

        If lItemPresent Then

            Select Case lItem.Class

                Case Outlook.OlObjectClass.olMail  ' only process MailItems

                    Try
                        lMailItem = lItem

                        Select Case lMailItem.MessageClass
                            Case "IPM.Note"
                                With lEmailMessage
                                    .MailItem = lMailItem
                                    .FromEmailAddress = lMailItem.SenderEmailAddress
                                    .SubjectLine = lMailItem.Subject
                                End With

                            Case "IPM.Note.Rules.OofTemplate.Microsoft"
                                ' DO NOTHING, let the Outlook rule "Automatic Reply" process this
                                Exit Select
                        End Select

                    Catch ex As Exception
                        LogMessage("*** ERROR *** GetMessage.olMail: " & ex.ToString)
                        lEmailMessage = Nothing ' just in case it is partially initialized

                    End Try

                Case Outlook.OlObjectClass.olReport
                    ' DO NOTHING, let Outlook rule "Undeliverable" process this
                    Exit Select

                Case Else

                    Try
                        ' TODO: not sure what else to do. . .
                        lItem.Move(gReportItemsFolder)

                    Catch ex As Exception
                        LogMessage("*** ERROR *** GetMessage.else: " & ex.ToString)

                    End Try

            End Select

        End If

        GetMessage = lEmailMessage

        lMailItem = Nothing
        lReportItem = Nothing
        lItem = Nothing
        lItems = Nothing

    End Function


    Public Sub InitFolders()

        Try
            gRouteTextFolder = gInbox.Folders(gRouteTextFolderName) ' folder must be INSIDE Inbox folder
        Catch
            LogMessage("*** ERROR *** InitFolders: Could not find gRouteTextFolderName; using Junk")
            gRouteTextFolder = gJunk
        End Try

        Try
            gNonRouteTextFolder = gInbox.Folders(gNonRouteTextFolderName) ' folder must be INSIDE Inbox folder
        Catch
            LogMessage("*** ERROR *** InitFolders: Could not find gNonRouteTextFolderName; using Junk")
            gRouteTextFolder = gJunk
        End Try

        Try
            gReportItemsFolder = gInbox.Folders(gReportItemsFolderName) ' folder must be INSIDE Inbox folder
        Catch
            LogMessage("*** ERROR *** InitFolders: Could not find gReportItemsFolderName; using Junk")
            gRouteTextFolder = gJunk
        End Try

    End Sub


    Public Sub FileMailItem(pMailItem As Outlook.MailItem, pRouteTextMessage As Boolean)

        Try
            If pRouteTextMessage Then
                pMailItem.Move(gRouteTextFolder)
            Else
                pMailItem.Move(gNonRouteTextFolder)
            End If

        Catch ex As Exception
            LogMessage("*** ERROR *** FileMailItem: Could not move message to a RouteText folder. " & ex.ToString)

            Try
                pMailItem.Delete() ' if we can't file it then just delete it
            Catch ex2 As Exception
                LogMessage("*** ERROR *** FileMailItem: Could not delete message. " & ex2.ToString)
            End Try
        End Try

    End Sub

End Module
