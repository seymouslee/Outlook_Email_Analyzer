Attribute VB_Name = "Module4"
Public filename As String
Public shortf As String
Public directory As String

Sub Extract2()
    Dim olItem As Outlook.MailItem, olMsg As Outlook.MailItem
    directory = CurDir() & "\2202Macro\"

    For Each olItem In Application.ActiveExplorer.Selection
        SaveEmailAttachmentsToFolder olItem, directory
    Next
    Set olMsg = Nothing
End Sub

Sub SaveEmailAttachmentsToFolder(olItem As Outlook.MailItem, DestFolder As String)

    Dim objAtt As Outlook.Attachment
    Dim i As Integer
    
    'If the directory doesn't exist, then create it
    If Dir(DestFolder, vbDirectory) = "" Then
        MkDir DestFolder
    End If
    
    'Check the email for attachments
    i = 0
    For Each Atmt In olItem.Attachments
        If LCase(Right(Atmt.filename, Len(ExtString))) = LCase(ExtString) Then
            shortf = Replace(Atmt.filename, " ", "_")
            filename = DestFolder & shortf
            Atmt.SaveAsFile filename
            'command shortf, CStr(i)
            cutil
            i = i + 1
        End If
    Next Atmt
    
    'done with extraction
    If i > 0 Then
        MsgBox "You can find the files here : " _
             & DestFolder, vbInformation, "Finished!"
    Else
        MsgBox "No attached files in your mail.", vbInformation, "Finished!"
    End If

End Sub

Sub cutil()
    Dim c As String
    c = "certutil -encode " & filename & " " & directory & "output.txt"
    Call Shell("cmd.exe /S /K" & c, vbHide)
End Sub
