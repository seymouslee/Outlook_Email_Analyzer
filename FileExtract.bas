Attribute VB_Name = "Module2"
Public filename As String
Public shortf As String

Sub ExtractAttachmentsV2()
    
    SaveEmailAttachmentsToFolder "test", "", "C:\Users\Mabel Lim\Documents"
    'command (shortf)

End Sub
Sub SaveEmailAttachmentsToFolder(OutlookFolderInInbox As String, _
                                 ExtString As String, DestFolder As String)
    Dim ns As NameSpace
    Dim Inbox As MAPIFolder
    Dim Item As Object
    Dim Atmt As Attachment
    'Dim filename As String
    Dim MyDocPath As String
    Dim i As Integer
    Dim wsh As Object
    Dim fs As Object

    On Error GoTo ThisMacro_err

    Set ns = GetNamespace("MAPI")
    Set Inbox = ns.GetDefaultFolder(olFolderInbox)

    If Right(DestFolder, 1) <> "\" Then
        DestFolder = DestFolder & "\"
    End If

    ' Check each message for attachments and extensions
    i = 0
    For Each Item In Inbox.Items
        For Each Atmt In Item.Attachments
            If LCase(Right(Atmt.filename, Len(ExtString))) = LCase(ExtString) Then
                'filename = DestFolder & Item.SenderName & " " & Atmt.filename
                shortf = Replace(Atmt.filename, " ", "_")
                filename = DestFolder & " " & shortf
                Atmt.SaveAsFile filename
                command shortf, CStr(i)
                i = i + 1
            End If
        Next Atmt
    Next Item

    ' Show this message when Finished
    If i > 0 Then
        MsgBox "You can find the files here : " _
             & DestFolder, vbInformation, "Finished!"
    Else
        MsgBox "No attached files in your mail.", vbInformation, "Finished!"
    End If

    ' Clear memory
ThisMacro_exit:
    Set Inbox = Nothing
    Set ns = Nothing
    Set fs = Nothing
    Set wsh = Nothing
    Exit Sub


End Sub


Sub command(fname As String, outputNo As String)
    Dim c As String
    'c = "certutil.exe -encode " + fname + " output.txt"
    
    'it works if I put the file directly into the documents folder
    'need to make sure output.txt does not exist before running this command
    'somehow cant encode the pdf file
    'c = "certutil.exe -encode lazy.txt output.txt"
    'certutil.exe -encode C:\Users\User\Documents\2202\Banana_Cake.pdf C:\Users\User\Documents\2202\encodestr.txt
    c = "certutil.exe -encode " + fname + " output" + outputNo + ".txt"
    MsgBox c
    Call Shell("cmd.exe /S /K" & c, vbNormalFocus)
End Sub

