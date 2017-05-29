NORM_VAL = 5000
PREV_CUT_OFF = 28
NEXT_CUT_OFF = 28
DATE_NOW = NOW()
SPLIT_MSEC = 1000

IF(DAY(DATE_NOW) < PREV_CUT_OFF) THEN
	MONTH_THIS = CStr(MONTH(DATE_NOW)-1)
	MONTH_NEXT = CStr(MONTH(DATE_NOW))
	TIME_THIS = " 00:00:00"
	TIME_NEXT = " 00:00:00" '& HOUR(DATE_NOW) & ":" & MINUTE(DATE_NOW) & ":" & SECOND(DATE_NOW)
ELSE
	MONTH_THIS = CStr(MONTH(DATE_NOW))
	MONTH_NEXT = CStr(MONTH(DATE_NOW)+1)
	TIME_THIS = " 00:00:00"
	TIME_NEXT = " 00:00:00"
END IF

DATE_THIS = CDate(MONTH_THIS & "/" & PREV_CUT_OFF & "/" & YEAR(DATE_NOW) & TIME_THIS)
DATE_NEXT = CDate(MONTH_NEXT & "/" & NEXT_CUT_OFF & "/" & YEAR(DATE_NOW) & TIME_NEXT)

'wscript.echo DateDiff("s",DATE_THIS, DATE_NEXT)
SEC_TOTAL   = DateDiff("s", DATE_THIS, DATE_NEXT)
SEC_ELAPSED = DateDiff("s", DATE_THIS, DATE_NOW)

Do
PERCENTAGE = (CLng(SEC_ELAPSED)/CLng(SEC_TOTAL)*1.0)
'RATE      = (60*60*24)/CLng(SEC_TOTAL)
FINAL_STR = CStr(PERCENTAGE*NORM_VAL)
FINAL_STR2 = Left(FINAL_STR, InStr(FINAL_STR,".")+5)
wscript.STDOUT.write _
"    " & DateDiff("s", "01/01/1970 08:00:00", Now()) & vbcrlf & _
Left(FINAL_STR, InStr(FINAL_STR,".")-1) & " DOLLAR " & _
MID(FINAL_STR, InStr(FINAL_STR,".")+1,2) & " SEN " & _
MID(FINAL_STR, InStr(FINAL_STR,".")+3,5) & "" '& vbcr

Wscript.Sleep SPLIT_MSEC
SEC_ELAPSED = SEC_ELAPSED + (SPLIT_MSEC/1000.0)
Loop

wscript.quit

Set colItems = objWMIService.ExecQuery("Select * From Win32_LocalTime")
 
For Each objItem in colItems
    strTime = objItem.Hour & ":" & objItem.Minute & ":" & objItem.Second
    dtmTime = CDate(strTime)
    Wscript.Echo FormatDateTime(dtmTime, vbFormatLongTime)
Next
