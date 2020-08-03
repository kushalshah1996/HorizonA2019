Function SystemVariable()
	Dim vDay
	Dim vMonth
	Dim vYear
	Dim vHour
	Dim vMinute
	Dim vSecond
	Dim vTime
	Dim vDate
	Dim vDateTime
	Dim vUserProfile
	
	vDateTime      = now()
	'vDateTime      = CDate("5-6-2022 8:4:3")
	'===Day of the Month===
	vDay           = day(vDateTime)
	if len(vDay)   = 1 then
		vDay = "0" & vDay
	End if
	
	'===Month of the Year===
	vMonth  = month(vDateTime)
	if len(vMonth) = 1 then
		vMonth     = "0" & vMonth
	End if
	
	'===Last two digits of Year===
	vYear          = year(vDateTime)
	vYear          = Right(vYear,2)
	
	'===Hours===
	vHour   = hour(vDateTime)
	if len(vHour)  = 1 then
		vHour      = "0" & vHour
	End if
	
	'===Minutes===
	vMinute = minute(vDateTime)
	if len(vMinute) = 1 then
		vMinute     = "0" & vMinute
	End if
	
	'===Seconds===
	vSecond = second(vDateTime)
	if len(vSecond) = 1 then
		vSecond     = "0" & vSecond
	End if
	
	vTime           = vHour & ":" & vMinute & ":" & vSecond
	vDate           = vDay & ":" & vMonth & ":" & vYear
	
	Set objShell    = CreateObject("WScript.Shell")
	vUserProfile    = objShell.ExpandEnvironmentStrings("%UserProfile%")
	
	SystemVariable  = vDate & "|" & vTime & "|" & vUserProfile & "\Documents\A2019"
	
	WScript.StdOut.WriteLine (SystemVariable)
End Function
