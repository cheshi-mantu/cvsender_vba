Attribute VB_Name = "cvsender"
Sub sendEmail(strRecipient As String)
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
Sub walker()

Dim lnLastLine As Long
Dim lnCounter As Long
Dim strTo, strToName, strSubject, strBodyPart As String
Dim xlActiveSheet As Worksheet
    
    Set xlActiveSheet = Application.ActiveSheet
    Debug.Print xlActiveSheet.Name
    'detect last filled line (row) in current (active) sheet
    lnLastLine = xlActiveSheet.Range("F1048576").End(xlUp).Row
    Debug.Print lnLastLine
    'Debug.Print "last used line in " & xlActiveSheet.Name & " is " & lnLastLine
    'Couter starts from 3rd line
    For lnCounter = 1 To lnLastLine
        If InStr(xlActiveSheet.Cells(lnCounter, 6).Text, "@") > 1 Then
            sendEmail xlActiveSheet.Cells(lnCounter, 6).Text
            Debug.Print xlActiveSheet.Cells(lnCounter, 6).Text
        End If
    Next
End Sub


