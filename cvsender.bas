Attribute VB_Name = "cvsender"
Private Sub sendEmail(strRecipient As String)
    On Error Resume Next
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(olMailItem)
    Set olAccount = oApp.Account
    Set olAccountTemp = oApp.Account
    Dim foundAccount As Boolean
    Dim strFrom As String
    Dim strHTMLBody
    
    strFrom = Application.Range("MAIL_ACCOUNT").Text
    strHTMLBody = Application.Range("HEADER_MSG").Text + Application.Range("BODY_MSG").Text + Application.Range("FOOTER_MSG").Text
    foundAccount = False
    Set olAccounts = oApp.Application.Session.Accounts
    For Each olAccountTemp In olAccounts
        'Debug.Print olAccountTemp.smtpAddress
        If InStr(olAccountTemp.smtpAddress, strFrom) > 0 Then
            Set olAccount = olAccountTemp
            foundAccount = True
            Exit For
        End If
    Next

    If foundAccount Then
        With oMail
            .To = strRecipient
            .HTMLBody = strHTMLBody
            .Subject = Application.Range("MSG_SUBJECT").Text
            Set .sendusingaccount = olAccount
            .Attachments.Add (Application.Range("CV_PATH").Text)
            '.Display
            '.Send
            .Save
        End With
    End If
    
    Set oApp = Nothing
    Set oMail = Nothing
    Set olAccounts = Nothing
    Set olAccount = Nothing
    Set olAccountTemp = Nothing
End Sub
Public Sub walker()
Dim lnLastLine As Long
Dim lnCounter As Long
Dim strTo, strToName, strSubject, strBodyPart As String
Dim xlSheet As Worksheet
Dim lngColnNum
Dim tmpArr
    'Set xlSheet = Application.ActiveSheet
    Set xlSheet = Worksheets(Application.Range("CELL_CONTACTS").Text)
    lngColnNum = findColWithEmails("email", Application.Range("CELL_CONTACTS").Text)
    
    'detect last filled line (row) in current (active) sheet
    lnLastLine = xlSheet.Cells(xlSheet.Rows.Count, lngColnNum).End(xlUp).Row
    'walk through found column from 1 to last row
    frmProgress.Show
    For lnCounter = 1 To lnLastLine
        If InStr(xlSheet.Cells(lnCounter, lngColnNum).Text, "@") > 1 Then
            frmProgress.txtboxProgress.Text = "Sending to " & xlSheet.Cells(lnCounter, lngColnNum).Text
            frmProgress.Repaint
            sendEmail xlSheet.Cells(lnCounter, lngColnNum).Text
            'Debug.Print xlSheet.Cells(lnCounter, lngColnNum).Text
        End If
    Next
End Sub
Public Sub checkMessage()
'define section
'==============================================
    Dim strFromAcc As String
    Dim strMessge As String
    Dim strCVFile As String
    Dim strHTMLHeader As String
    Dim strHTMLFooter As String
    Dim envVar As String 'store env variable value here
    Dim objFSys 'define file system object
    Dim objFile ' define file object
    Dim strFileName As String
    Dim objShell 'define shell scripting object
    Dim oApp As Object
    Dim olAccount
    Dim olAccountTemp
    Dim strCVFileNotFound
    Dim strBrowserPath As String
    Dim oNS As Object 'define outlook namespace
'Init section
'===============================================
 On Error Resume Next
    strCVFileNotFound = " CV file not found by the way"
    envVar = CStr(Environ("TEMP")) 'we'll use it to build path for checker html file
    strFileName = "checking_message.html"
    'Set oApp = CreateObject("Outlook.Application")
    'Set olAccount = oApp.Account
    'Set olAccountTemp = oApp.Account
    Set objFSys = CreateObject("Scripting.FilesystemObject") 'init file system object
    strFileName = envVar + "\" + strFileName
    Debug.Print strFileName
    Set objFile = objFSys.CreateTextFile(strFileName)
    Set objFile = objFSys.getFile(strFileName)
    Set objFile = objFSys.OpenTextFile(strFileName, 2, 1, -2)
    
    objFile.Write ("<H1>Message subject</H1>")
    objFile.Write (Application.Range("MSG_SUBJECT").Text)
    
    objFile.Write ("<H1>Message body with the signature</H1>")
    objFile.Write (Application.Range("HTML_HEADER").Text + Application.Range("HEADER_MSG").Text + Application.Range("BODY_MSG").Text + Application.Range("FOOTER_MSG").Text + Application.Range("HTML_FOOTER").Text)
    
    objFile.Write ("<H1>Path to the file with your CV</H1>")
    strBrowserPath = Replace(Application.Range("CV_PATH").Text, "\", "/")
    
    Debug.Print strBrowserPath
    Debug.Print objFileSys.FileExists(Application.Range("CV_PATH").Text)
    
    If objFSys.FileExists(Application.Range("CV_PATH").Text) Then
        strCVFileNotFound = ""
    End If
    
    
    objFile.Write ("<a href='file:///" + strBrowserPath + "'" + "target='_blank'>Check Your CV presence by following this link</a>" + strCVFileNotFound)
    objFile.Write ("<H1>emails will be sent from this address</H1>")
    objFile.Write (getAccountName)
    
    objFile.Close
    
    Set objShell = CreateObject("Shell.Application")
    objShell.ShellExecute strFileName, "", "", "", 1
    
    
'Destructs
'===============================================
    Set oApp = Nothing
    Set olAccount = Nothing
    Set olAccountTemp = Nothing
    Set objFSys = Nothing
    Set objShell = Nothing
End Sub
Private Function getAccountName()
    On Error Resume Next
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(olMailItem)
    Set olAccount = oApp.Account
    Set olAccountTemp = oApp.Account
    Dim foundAccount As Boolean
    Dim strFrom As String
    Dim strHTMLBody
    
    strFrom = Application.Range("MAIL_ACCOUNT").Text
    foundAccount = False
    Set olAccounts = oApp.Application.Session.Accounts
    For Each olAccountTemp In olAccounts
        'Debug.Print olAccountTemp.smtpAddress
        If InStr(olAccountTemp.smtpAddress, strFrom) > 0 Then
            Set olAccount = olAccountTemp
            foundAccount = True
            Exit For
        End If
    Next

    If foundAccount Then
        getAccountName = olAccount.smtpAddress
    'Debug.Print olAccount.smtpAddress
    Else
        getAccountName = "No such account"
    End If
    
    Set oApp = Nothing
    Set oMail = Nothing
    Set olAccounts = Nothing
    Set olAccount = Nothing
    Set olAccountTemp = Nothing
    
End Function
Private Function findColWithEmails(strFindWord As String, strWSName As String)
    Dim ws As Worksheet
    Dim rngEmailCol As Range
    Set ws = Worksheets(strWSName)
    Set rngEmailCol = ws.Cells.Find(What:="email", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        findColWithEmails = rngEmailCol.Column
    Set ws = Nothing
    Set rngEmailCol = Nothing
End Function

