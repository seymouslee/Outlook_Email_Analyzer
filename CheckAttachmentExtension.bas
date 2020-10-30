Attribute VB_Name = "Module4"
Public filename As String
Public shortf As String
Public directory As String
Public checkExtension As Boolean
Public checkExtensionAll As Boolean

Sub Extract()
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
    Dim goExplorer As Integer

    
    'If the directory doesn't exist, then create it
    If Dir(DestFolder, vbDirectory) = "" Then
        MkDir DestFolder
    End If
    
    'creation of quarantine folder
    If Dir(CurDir() & "\2202Quarantine\", vbDirectory) = "" Then
        MkDir CurDir() & "\2202Quarantine\"
    End If
    
    'Check the email for attachments
    i = 0
    checkExtension = False
    checkExtensionAll = False
    For Each Atmt In olItem.Attachments
        If LCase(Right(Atmt.filename, Len(ExtString))) = LCase(ExtString) Then
            shortf = Replace(Atmt.filename, " ", "_")
            filename = DestFolder & shortf
            Atmt.SaveAsFile filename
            cutil                           'run certUtil to find the certificate for the file
            Wait (1)                        'give time for certUtil in cmd to execute before moving on
            readEx                          'reading the certificate for the file and check if extension is potentially malicious
            If checkExtension = True Then   'if the extension turns out to be potentially malicious
                moveToQuarantine (shortf)   'move the certification to another folder where user can access later on
                deletefiles
            End If
            i = i + 1
            checkExtension = False
        End If
    Next Atmt
    
    deletefiles
    
    If checkExtensionAll = True Then
        MsgBox "Warning: attachment(s) in this email may be malicious", vbExclamation
        goExplorer = MsgBox(Prompt:="Certification file and header report can be found in " & CurDir() & "\2202Quarantine\" & _
        vbNewLine & vbNewLine & "Do you want to open directory in file explorer?", Buttons:=vbOKCancel)
        If goExplorer = vbOK Then
            'if the user wants to see the file, the macro will open the directory
            Shell "cmd /c start """" explorer.exe " & CurDir() & "\2202Quarantine\", vbHide
        End If
    End If

End Sub

Sub cutil()
    Dim c As String
    c = "certutil -encode " & filename & " " & directory & "output.txt"
    Call Shell("cmd.exe /S /K" & c, vbHide)
End Sub

Sub deletefiles()
    On Error Resume Next
    Kill (directory & "*.*")
    On Error GoTo 0
End Sub

Sub moveToQuarantine(filename As String)
    Dim FSO As Object
    Dim sourceD As String, destinationD As String

    Set FSO = CreateObject("Scripting.Filesystemobject")
    sourceD = directory & "output.txt"
    destinationD = CurDir() & "\2202Quarantine\cert_" & filename

    FSO.MoveFile source:=sourceD, destination:=destinationD
End Sub

Sub readEx()

Dim FSO, FileIn, txtLine
Dim fdirectory As String
fdirectory = directory & "output.txt"

'executable file extensions in base64
extenstions = Array("TVo", "XyeoiQ", "yv66vg", "QkxJMjIzUQ", "HX0", "183Gmg", "UEsDBBQA")

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FileIn = FSO.OpenTextFile(fdirectory, 1)

Do Until FileIn.AtEndOfStream
    txtLine = FileIn.ReadLine
    If Len(txtLine) > 0 Then
        For i = 0 To UBound(extenstions)
            If InStr(1, txtLine, extenstions(i), vbTextCompare) > 0 Then
                    checkExtension = True
                    checkExtensionAll = True
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

