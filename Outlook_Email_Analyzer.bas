Attribute VB_Name = "Module4"
Public filename, fn As String
Public shortf As String
Public directory As String
Public checkExtension As Boolean
Public checkExtensionAll As Boolean
Public errorlog() As String
Public errorEx, allerrors As String

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
    Dim i, x As Integer
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
            fn = Atmt.filename
            shortf = Replace(Atmt.filename, " ", "_")
            filename = DestFolder & shortf
            Atmt.SaveAsFile filename
            cutil                           'run certUtil to find the certificate for the file
            Wait (1)                        'give time for certUtil in cmd to execute before moving on
            readEx                          'reading the certificate for the file and check if extension is potentially malicious
            If checkExtension = True Then   'if the extension turns out to be potentially malicious
                moveToQuarantine (shortf)   'move the certification to another folder where user can access later on
                deletefiles
                ReDim Preserve errorlog(i)
                errorlog(i) = shortf & " may actually be " & errorEx & " file"  'for error logging
            End If
            i = i + 1
            checkExtension = False
        End If
    Next Atmt
    
    deletefiles
    
    x = 0
    Do While x <= i - 1
        allerrors = allerrors & errorlog(x) & Chr(10)
        x = x + 1
    Loop
    
    If checkExtensionAll = True Then
        ViewInternetHeader
        MsgBox "Warning: attachment(s) in this email may be malicious" & Chr(10) & Chr(10) & allerrors, vbExclamation
        goExplorer = MsgBox(Prompt:="Certification file and header report can be found in " & CurDir() & "\2202Quarantine\" & _
        vbNewLine & vbNewLine & "Do you want to open directory in file explorer?", Buttons:=vbOKCancel)
        If goExplorer = vbOK Then
            'if the user wants to see the file, the macro will open the directory
            Shell "cmd /c start """" explorer.exe " & CurDir() & "\2202Quarantine\", vbHide
        End If
    Else
        MsgBox "Scan complete: nothing wrong with the attachments"
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
Extensions = Array("TVo", "XyeoiQ", "yv66vg", "QkxJMjIzUQ", "HX0", "183Gmg", "UEsDBBQA")
exname = Array(".exe (executable)", ".jar (JavaScript source code script)", ".jar (JavaScript source code script)", ".bin (Binary executable)", ".ws (Microsoft Windows script)", ".wmf (Windows Metafile Format)", ".jar (JavaScript source code script)")

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FileIn = FSO.OpenTextFile(fdirectory, 1)

Do Until FileIn.AtEndOfStream
    txtLine = FileIn.ReadLine
    If Len(txtLine) > 0 Then
        For i = 0 To UBound(Extensions)
            If InStr(1, txtLine, Extensions(i), vbTextCompare) > 0 Then
                    checkExtension = True
                    checkExtensionAll = True
                    errorEx = exname(i)
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


Sub ViewInternetHeader()
      Dim objMail As Outlook.MailItem
    Dim objPropertyAccessor As Outlook.PropertyAccessor
    Dim strHeader As String
    Dim strTempFolder As String
    Dim objFileSystem As Object
    Dim strTextFile As String
    Dim objTextFile As Object
    Dim objForward As Outlook.MailItem
    
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    Dim rCount As Long
    Dim bXStarted As Boolean
    Dim enviro As String
    Dim strPath As String
    
    Dim currentExplorer As Explorer
    Dim Selection As Selection
    Dim olItem, olMsg As Outlook.MailItem
    Dim obj As Object
    Dim strColA, strColB, strColC, strColD, strColE, strColF, strColG, strColH, strColI As String
 
 
               
' Get Excel set up
     On Error Resume Next
     Set xlApp = GetObject(, "Excel.Application")
     If Err <> 0 Then
         Application.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("Excel.Application")
         bXStarted = True
     End If
     On Error GoTo 0
 
    Select Case Outlook.Application.ActiveWindow.Class
           Case olInspector
                Set objMail = ActiveInspector.CurrentItem
           Case olExplorer
                Set objMail = ActiveExplorer.Selection.Item(1)
    End Select
 
    'Get the Internet Headers
    Set objPropertyAccessor = objMail.PropertyAccessor
    strHeader = objPropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
    
    'Mail Authentication
    Dim extractedSPF As String: extractedSPF = InStr(strHeader, "spf=pass")
    Dim extractedDKIM As String: extractedDKIM = InStr(strHeader, "dkim=pass")
    Dim extractedDMARC As String: extractedDMARC = InStr(strHeader, "dmarc=pass")
    Dim passMailAuthentication As String
       If extractedSPF > 0 And extractedDKIM > 0 And extractedDMARC > 0 Then
        passMailAuthentication = "Email Authenticated"
       Else
        passMailAuthentication = "Email Not Authenticated"
       End If
       
    'Get IP Addresses Function
    Dim ipAddrs() As String
    Dim ipAddrResults As String
    Dim requestArray() As String
    Dim responseString As String
    ipAddrs = getIPAddresses(strHeader)
    If Len(ipAddrs(0)) > 0 Then
     Dim oRequest As Object
        Set oRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        oRequest.Open "GET", "https://api.ipregistry.co/" & ipAddrs(0) & "?key=fvenf3i7s06ri0"
        oRequest.Send
        Dim requestString As String: requestString = oRequest.responseText
        requestArray = Split(requestString, ",") 'split response text
        responseString = Join(requestArray, vbCrLf) 'break into new line
       ipAddrResults = Join(ipAddrs, vbCrLf)
    End If
    
    Set xlWB = xlApp.Workbooks.Add
Set xlSheet = xlWB.Sheets("Sheet1")
'## end use new workbook

' Add column names
  xlSheet.Range("A1") = "Sender"
  xlSheet.Range("B1") = "Sender Address"
  xlSheet.Range("C1") = "Message Body"
  xlSheet.Range("D1") = "Sent To"
  xlSheet.Range("E1") = "Recieved Time"
  xlSheet.Range("F1") = "Mail-Authentication"
  xlSheet.Range("G1") = "IP Addresses"
  xlSheet.Range("H1") = "Curl IP Addresses"
  xlSheet.Range("I1") = "Internet Headers"

' Process the message record
    
  On Error Resume Next
'Find the next empty line of the worksheet
rCount = xlSheet.Range("A" & xlSheet.Rows.Count).End(-4162).Row
'needed for Exchange 2016. Remove if causing blank lines.
rCount = rCount + 1

' get the values from outlook
Set currentExplorer = Application.ActiveExplorer
Set Selection = currentExplorer.Selection
  For Each obj In Selection

    Set olItem = obj
    
 'collect the fields
    strColA = olItem.SenderName
    strColB = olItem.SenderEmailAddress
    strColC = olItem.Body
    strColD = olItem.To
    strColE = olItem.ReceivedTime
    strColF = passMailAuthentication
    strColG = ipAddrResults
    strColH = responseString
    strColI = strHeader
    
    
    
'### Get all recipient addresses
' instead of To names
Dim strRecipients As String
Dim Recipient As Outlook.Recipient
For Each Recipient In olItem.Recipients
 strRecipients = Recipient.Address & "; " & strRecipients
 Next Recipient

  strColD = strRecipients
'### end all recipients addresses

'### Get the Exchange address
' if not using Exchange, this block can be removed
 Dim olEU As Outlook.ExchangeUser
 Dim oEDL As Outlook.ExchangeDistributionList
 Dim recip As Outlook.Recipient
 Set recip = Application.Session.CreateRecipient(strColB)

If InStr(1, strColB, "/") > 0 Then
' if exchange, get smtp address
    Select Case recip.AddressEntry.AddressEntryUserType
       Case OlAddressEntryUserType.olExchangeUserAddressEntry
         Set olEU = recip.AddressEntry.GetExchangeUser
         If Not (olEU Is Nothing) Then
             strColB = olEU.PrimarySmtpAddress
         End If
       Case OlAddressEntryUserType.olOutlookContactAddressEntry
         Set olEU = recip.AddressEntry.GetExchangeUser
         If Not (olEU Is Nothing) Then
            strColB = olEU.PrimarySmtpAddress
         End If
       Case OlAddressEntryUserType.olExchangeDistributionListAddressEntry
         Set oEDL = recip.AddressEntry.GetExchangeDistributionList
         If Not (oEDL Is Nothing) Then
            strColB = olEU.PrimarySmtpAddress
         End If
     End Select
End If
' ### End Exchange section

'write them in the excel sheet
  xlSheet.Range("A" & rCount) = strColA ' sender name
  xlSheet.Range("B" & rCount) = strColB ' sender address
  xlSheet.Range("C" & rCount) = strColC ' message body
  xlSheet.Range("D" & rCount) = strColD ' sent to
  xlSheet.Range("E" & rCount) = strColE ' recieved time
  xlSheet.Range("F" & rCount) = strColF ' mail authentication
  xlSheet.Range("G" & rCount) = strColG ' collect all ip address
  xlSheet.Range("H" & rCount) = strColH ' collects api
  xlSheet.Range("I" & rCount) = strColI 'collect internet headers
  
 
'Next row
  rCount = rCount + 1

' size the cells
    xlSheet.Columns("A:E").EntireColumn.AutoFit
    xlSheet.Columns("C:C").ColumnWidth = 100
    xlSheet.Columns("D:D").ColumnWidth = 30
    xlSheet.Columns("F:F").ColumnWidth = 50
    xlSheet.Columns("G:G").ColumnWidth = 40
    xlSheet.Columns("H:H").ColumnWidth = 30
    xlSheet.Columns("I:I").ColumnWidth = 50
    xlSheet.Range("A2").Select
    xlSheet.Columns("A:I").VerticalAlignment = xlTop

 Next

MsgBox "saved at " & CurDir() & "\2202Quarantine\emailheader.xlsx"
xlWB.SaveCopyAs filename:=CurDir() & "\2202Quarantine\emailheader.xlsx"


' to save but not close
'xlWB.Save
    
    
End Sub

Function getIPAddresses(ByVal MsgHeader As String) As String()
    Dim tempArr() As String, i As Long, regEx As Object, regC As Object
    Set regEx = CreateObject("vbscript.regexp")
    ReDim tempArr(0)
    With regEx
        .Global = True
        .MultiLine = True
        .pattern = "\[?(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\]?"
    End With
    
    If regEx.test(MsgHeader) Then
        Set regC = regEx.Execute(MsgHeader)
        ReDim tempArr(regC.Count - 1)
        For i = 0 To regC.Count - 1
            tempArr(i) = regC.Item(i).SubMatches(0)
        Next
    End If
    Set regEx = Nothing
    Set regC = Nothing
    getIPAddresses = tempArr
            
End Function
