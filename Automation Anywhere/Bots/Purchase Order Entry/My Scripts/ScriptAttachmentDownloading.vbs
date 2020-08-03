Dim vSavefolder
Dim vAllEmail
Dim vValidEmail
Dim vBotEmailID
Dim vOutlookAccount
Dim vSubject
Dim vCountAttachment
Dim vTimeReceived
Dim vTimeStamp
Dim vAttachmentFileName
Dim vAttachmentExtension
Dim vFileDownloadPath

'==Logic to Assign Required Variables Begins==
vSavefolder        = WScript.Arguments.Item(0)
'vSavefolder       = "C:\Users\Parthiban.Nadar\Documents\A2019\Sales Order Entry\Current Folder\Email Attachment Folder"

vValidEmailAddress = WScript.Arguments.Item(1)
vTextFile          = vSavefolder & "\Email Report.txt"
'==Logic to Assign Required Variables Ends==
Set fso            = CreateObject("Scripting.FileSystemObject")

'==Logic to Open an outlook application if does not exist Begins==

On Error Resume Next
Set olApp = GetObject(, "Outlook.Application")
If Err.Number <> 0 Then    'Could not get instance of Outlook, so create a new one
Err.Clear
Set olApp = CreateObject("Outlook.Application")
End If
on error goto 0
Set olns = olApp.GetNameSpace("MAPI")
olns.logon "Outlook",,False,True

'==Logic to Open an outlook application if does not exist Ends==

Set vInboxFolder    = olns.GetDefaultFolder(6) ' here it selects the inbox folder of account.
set MailItems       = vInboxFolder.Items
'==Loop through emails of Bot Email ID ONLY Begins==	

For i = 1 To MailItems.Count  
	If MailItems.Item(i).Unread = true then
		vSubject = Trim(Ucase(MailItems.Item(i).Subject))
		If MailItems.Item(i).SenderEmailType = "SMTP" Then
			vSenderEmailAddress = MailItems.Item(i).SenderEmailAddress
			vSenderEmailAddress = LCase(Trim(vSenderEmailAddress))
		ElseIf MailItems.Item(i).SenderEmailType = "EX" Then
			vSenderEmailAddress = MailItems.Item(i).Sender.GetExchangeUser.PrimarySmtpAddress
			vSenderEmailAddress = LCase(Trim(vSenderEmailAddress))
		End If
		'End of "MailItems.Item(i).SenderEmailType = "SMTP"" If Statement
		
		if InStr(vSenderEmailAddress,vValidEmailAddress) <> 0 Then
			fso.CreateTextFile vTextFile
			Set ts = fso.OpenTextFile(vTextFile, 8, True, 0)
			ts.WriteLine "Email From = "    & vSenderEmailAddress
			ts.WriteLine "Email Subject = " & vSubject
			vCountAttachment = MailItems.Item(i).Attachments.Count
			if vCountAttachment > 0 Then
				For j = 1 to vCountAttachment
					vTimeReceived = MailItems.Item(i).ReceivedTime
					vTimeStamp           = Year(vTimeReceived)                  & _ 
										Right("0" & Month(vTimeReceived),2)  & _ 
										Right("0" & Day(vTimeReceived),2)    & _ 
										Right("0" & Hour(vTimeReceived),2)   & _ 
										Right("0" & Minute(vTimeReceived),2) & _
										Right("0" & Second(vTimeReceived),2)
				
					ts.WriteLine "Attachment Found = " & "Yes " & vTimeStamp
					vAttachmentFileName  = MailItems.Item(i).Attachments(j).FileName
					vFileDownloadPath    = vSavefolder & "\" & vAttachmentFileName
					MailItems.Item(i).Attachments(j).SaveAsFile vFileDownloadPath
				Next
				'For "j = 1 to vCountAttachment" Loop
			End if
			'End of "vCountAttachment > 0" If Statement
			ts.Close
			MailItems.Item(i).UnRead = False
			Exit For
			'Exiting the mails so that the remaining mails are unread
		End If
		'Exiting If vSenderEmail includes vValidEmailAddress
	End if
	'End of mail = Unread If Statement
Next
'==Loop through emails of Bot Email ID ONLY Ends==

