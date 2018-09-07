﻿Imports Outlook = Microsoft.Office.Interop.Outlook
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

                        If lMailItem.MessageClass = "IPM.Note" Then
                            With lEmailMessage
                                .MailItem = lMailItem
                                .FromEmailAddress = lMailItem.SenderEmailAddress
                                .SubjectLine = lMailItem.Subject
                            End With
                        End If

                    Catch ex As Exception
                        LogMessage("*** ERROR *** GetMessage.olMail: " & ex.ToString)
                        lEmailMessage = Nothing ' just in case it is partially initialized

                    End Try

                Case Outlook.OlObjectClass.olReport  ' just file ReportItems for now; an Outlook rule might have caught this anyway

                    ' DO NOTHING, let Outlook rule "Undeliverable" forward and file this item
                    Exit Select

                    'Try
                    '    lReportItem = lItems.Item(1)
                    '    lReportItem.Move(gReportItemsFolder)
                    'Catch
                    '    LogMessage("*** ERROR *** GetMessage: Could not move ReportItem to folder; deleting it")
                    '    lItems.Item(1).Delete()
                    'End Try

                    ' SAVE: Possible code snippet for parsing HTML Outlook report to extract original email address
                    'lReportString = Text.Encoding.ASCII.GetString(Text.Encoding.Unicode.GetBytes(lItems.Item(1).Body))
                    'lStartPosition = InStr(lReportString, "To: ") + 4
                    'lEndPosition = InStr(lStartPosition, lReportString, vbCrLf) - lStartPosition
                    'lOriginalRecipient = Mid(lReportString, lStartPosition, lEndPosition)

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
