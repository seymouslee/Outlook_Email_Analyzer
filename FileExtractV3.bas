Attribute VB_Name = "Module4"
Public FileName As String
Public shortf As String
Public directory As String
Public checkExtension As Boolean

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
    Dim ex As String
    Dim started As Single: started = Timer

    
    'If the directory doesn't exist, then create it
    If Dir(DestFolder, vbDirectory) = "" Then
        MkDir DestFolder
    End If
    
    'Check the email for attachments
    i = 0
    checkExtension = False
    For Each Atmt In olItem.Attachments
        If LCase(Right(Atmt.FileName, Len(ExtString))) = LCase(ExtString) Then
            shortf = Replace(Atmt.FileName, " ", "_")
            FileName = DestFolder & shortf
            Atmt.SaveAsFile FileName
            cutil
            Wait (1)
            readEx
            If checkExtension = True Then
                MsgBox "Warning: attachment(s) in this email may be malicious", vbExclamation
            End If
            i = i + 1
        End If
    Next Atmt

End Sub

Sub cutil()
    Dim c As String
    c = "certutil -encode " & FileName & " " & directory & "output.txt"
    Call Shell("cmd.exe /S /K" & c, vbHide)
End Sub

Sub deletefiles()
    On Error Resume Next
    Kill (directory & "*.*")
    On Error GoTo 0
End Sub

Sub readEx()

Dim FSO, FileIn, strTmp
Dim fdirectory As String
fdirectory = directory & "output.txt"

Dim extenstions(6) As String
'extenstions = Array("TVo", "XyeoiQ", "yv66vg", "QkxJMjIzUQ", "HX0", "183Gmg")
arWords = Array("aGVs", "support", "xyz")

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FileIn = FSO.OpenTextFile(fdirectory, 1)

Do Until FileIn.AtEndOfStream
    strTmp = FileIn.ReadLine
    If Len(strTmp) > 0 Then
        For i = 0 To UBound(arWords)
            If InStr(1, strTmp, arWords(i), vbTextCompare) > 0 Then
                    checkExtension = True
                Exit For
            End If
        Next
    End If
Loop

FileIn.Close

End Sub

Sub Wait(seconds As Integer)
  Dim now As Long
  now = Timer()
  Do
      DoEvents
  Loop While (Timer < now + seconds)
End Sub
