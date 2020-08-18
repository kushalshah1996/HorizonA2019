str = WScript.Arguments.Item(0)
'str = "CNWY"
str = Trim(str)
vOutput = str
for i = 1 to len(str)
	vLetter  = Mid(str,i,1)
	vBoolean = IsNumeric(vLetter)
	if (vBoolean = TRUE) Then
		vOutput = "Account Number"
		exit For
	End if
Next
if InStr(str,"FEDEX") <> 0 then	
	for i = 1 to len(str)
	vLetter  = Mid(str,i,1)
	vBoolean = IsNumeric(vLetter)
	if (vBoolean = TRUE) Then
		vOutput = "FEDEX - ACC"
		exit For
	End if
	Next
End if

Set networkInfo  = CreateObject("WScript.NetWork") 
vUserName        = networkInfo.UserName

Set fso   = CreateObject("Scripting.FileSystemObject")
vTextFile = "C:\Users\" + vUserName + "\Documents\A2019\Purchase Order Entry\Current Folder\ShipVia.txt"
fso.CreateTextFile vTextFile
Set ts  = fso.OpenTextFile(vTextFile, 8, True, 0)
ts.WriteLine vOutput
ts.Close
