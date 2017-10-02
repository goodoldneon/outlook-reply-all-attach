Public Class ThisAddIn
    Shared outlookApp As Outlook.Application

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        outlookApp = Me.Application
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Public Shared Sub ReplyWithAttachments()
        Dim rpl As Outlook.MailItem
        Dim itm As Object

        itm = GetCurrentItem()

        If Not itm Is Nothing Then
            rpl = itm.ReplyAll
            CopyAttachments(itm, rpl)
            rpl.Display()
        End If

        rpl = Nothing
        itm = Nothing
    End Sub

    Private Shared Function GetCurrentItem() As Object
        Select Case TypeName(outlookApp.ActiveWindow)
            Case "Explorer"
                GetCurrentItem = outlookApp.ActiveExplorer.Selection.Item(1)
            Case "Inspector"
                GetCurrentItem = outlookApp.ActiveInspector.CurrentItem
            Case Else
                MsgBox("ActiveWindow isn't 'Explorer' or 'Inspector'.")
                GetCurrentItem = Nothing
        End Select
    End Function

    Private Shared Sub CopyAttachments(objSourceItem, objTargetItem)
        Dim fso
        Dim fldTemp
        Dim strPath
        Dim strFile

        fso = CreateObject("Scripting.FileSystemObject")
        fldTemp = fso.GetSpecialFolder(2) ' Temporary folder
        strPath = fldTemp.Path & "\"

        For Each objAtt In objSourceItem.Attachments
            strFile = strPath & objAtt.FileName
            objAtt.SaveAsFile(strFile)
            objTargetItem.Attachments.Add(strFile, , , objAtt.DisplayName)
            fso.DeleteFile(strFile)
        Next

        fldTemp = Nothing
        fso = Nothing
    End Sub
End Class
