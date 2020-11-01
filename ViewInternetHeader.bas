Attribute VB_Name = "Module1"
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
                Set objMail = ActiveExplorer.Selection.item(1)
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
 xlApp.Visible = True

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
            tempArr(i) = regC.item(i).SubMatches(0)
        Next
    End If
    Set regEx = Nothing
    Set regC = Nothing
    getIPAddresses = tempArr
            
End Function



