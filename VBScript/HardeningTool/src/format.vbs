

FUNCTION WRITE_OUTPUT_5(BYVAL FIELD1, BYVAL FIELD2, BYVAL FIELD3, BYVAL FIELD4, BYVAL FIELD5)
	IF HTML_MODE = TRUE THEN
		wscript.stdout.write "<tr class='itemrow'>"
		wscript.stdout.write "<td class='field1'>" & FIELD1 & "</td>"
		wscript.stdout.write "<td class='field2'>" & FIELD2 & "</td>"
		wscript.stdout.write "<td class='field3'>" & FIELD3 & "</td>"
		wscript.stdout.write "<td class='field4'>" & FIELD4 & "</td>"
		wscript.stdout.write "<td class='field5'>" & FIELD5 & "</td>"
		wscript.stdout.write "</tr>"
		wscript.stdout.write vbCRLF
	ELSE
		wscript.stdout.write FPadding(FIELD1,  7) & " "
		wscript.stdout.write FPadding(FIELD2, 30) & " "
		wscript.stdout.write FPadding(FIELD3, 20) & " "
		wscript.stdout.write FPadding(FIELD4, 20)
		wscript.stdout.write FPadding(FIELD5, 20)
		wscript.stdout.write vbCRLF
	END IF
END FUNCTION

FUNCTION WRITE_OUTPUT(BYVAL FIELD1, BYVAL FIELD2, BYVAL FIELD3, BYVAL FIELD4)
	IF HTML_MODE = TRUE THEN
		wscript.stdout.write "<tr class='itemrow'>"
		wscript.stdout.write "<td class='field_first'></td>"
		wscript.stdout.write "<td class='field1'>" & FIELD1 & "</td>"
		wscript.stdout.write "<td class='field2'>" & FIELD2 & "</td>"
		wscript.stdout.write "<td class='field3'>" & FIELD3 & "</td>"
		wscript.stdout.write "<td class='field4'>" & FIELD4 & "</td>"
		wscript.stdout.write "<td class='field_last'></td>"
		wscript.stdout.write "</tr>"
		wscript.stdout.write vbCRLF
	ELSE
		wscript.stdout.write FPadding(FIELD1,  7) & " "
		wscript.stdout.write FPadding(FIELD2, 50) & " "
		wscript.stdout.write FPadding(FIELD3, 20) & " "
		wscript.stdout.write FPadding(FIELD4, 20)
		wscript.stdout.write vbCRLF
	END IF
END FUNCTION

FUNCTION WRITE_OUTPUT_SECTION_START(BYVAL HEADER, BYVAL DESCRIPTION)
	wscript.stderr.write FPadding(HEADER & "... ", 50)
	IF HTML_MODE = TRUE THEN HTML_WRITE_OUTPUT_SECTION_START HEADER, DESCRIPTION: EXIT FUNCTION
	wscript.stdout.write ""
END FUNCTION

FUNCTION WRITE_OUTPUT_SECTION_START_COL(BYVAL HEADER, BYVAL DESCRIPTION, BYVAL COL)
	wscript.stderr.write FPadding(HEADER & "... ", 50)
	IF HTML_MODE = TRUE THEN HTML_WRITE_OUTPUT_SECTION_START_COL HEADER, DESCRIPTION, COL: EXIT FUNCTION
	wscript.stdout.write ""
END FUNCTION

FUNCTION WRITE_OUTPUT_SECTION_END(BYVAL FOOTER)
	wscript.stderr.write "DONE." & vbcrlf
	IF HTML_MODE = TRUE THEN HTML_WRITE_OUTPUT_SECTION_END FOOTER: EXIT FUNCTION
	wscript.stdout.write vbcrlf & vbcrlf
END FUNCTION


FUNCTION HTML_WRITE_OUTPUT_SECTION_START(BYVAL HEADER, BYVAL DESCRIPTION)
	HTML_WRITE_OUTPUT_SECTION_START_COL HEADER, DESCRIPTION, 6
END FUNCTION

FUNCTION HTML_WRITE_OUTPUT_SECTION_START_COL(BYVAL HEADER, BYVAL DESCRIPTION, BYVAL COL)
	'wscript.stdout.write "<table class='sortable' cellspacing=1 cellpadding=0 id='" & Replace(HEADER," ","_") & "'><tr><th colspan=" & COL & ">" & HEADER & "</th></tr>"
	FINAL_NAME = HEADER
	FINAL_NAME = Replace(FINAL_NAME," ","_")
	FINAL_NAME = Replace(FINAL_NAME,".","_")
	FINAL_NAME = Replace(FINAL_NAME,"(","")
	FINAL_NAME = Replace(FINAL_NAME,")","")
	IF Instr(FINAL_NAME, "-") > 2 THEN FINAL_NAME = Left(FINAL_NAME, Instr(FINAL_NAME, "-") -2)

	wscript.stdout.write "<h1>" & HEADER & "</h1>"
	IF NOT DESCRIPTION = "" THEN wscript.stdout.write "<h2>" & DESCRIPTION & "</h2>"
	wscript.stdout.write "<table class='sortable' cellspacing=0 cellpadding=0 id='" & FINAL_NAME & "'>" '<tr><th colspan=" & COL & ">" & HEADER & "</th></tr>"
		
END FUNCTION

FUNCTION HTML_WRITE_OUTPUT_SECTION_END(BYVAL FOOTER)
	wscript.stdout.write "</table>" & vbcrlf & vbcrlf 
	wscript.stdout.write "<div style='clear: both; height: 30px'></div>"
END FUNCTION


FUNCTION CONFIRM_CHARACTER
	IF HTML_MODE = TRUE THEN
		CONFIRM_CHARACTER = "<img src='images/accept.gif' />"
		'CONFIRM_CHARACTER = "<img src='images/blank.png' class='icon success' />"
	ELSE
		CONFIRM_CHARACTER = "/"
	END IF
END FUNCTION


FUNCTION ERROR_CHARACTER
	IF HTML_MODE = TRUE THEN
		ERROR_CHARACTER = "<img src='images/error.gif' />"
		'ERROR_CHARACTER = "<img src='images/blank.png' class='icon error' />"
	ELSE
		ERROR_CHARACTER = "X (ERROR)"
	END IF
END FUNCTION


FUNCTION NA_CHARACTER
	IF HTML_MODE = TRUE THEN
		NA_CHARACTER = "<img src='images/na.gif' /> N/A"
		'NA_CHARACTER = "<img src='images/blank.png' class='icon na' /> N/A"
	ELSE
		NA_CHARACTER = "N/A"
	END IF
END FUNCTION

FUNCTION WARNING_CHARACTER
	IF HTML_MODE = TRUE THEN
		WARNING_CHARACTER = "<img src='images/warning.gif' />"
		'WARNING_CHARACTER = "<img src='images/blank.png' class='icon warning' /> N/A"
	ELSE
		WARNING_CHARACTER = "(WARNING)"
	END IF
END FUNCTION

FUNCTION MANUAL_CHARACTER
	IF HTML_MODE = TRUE THEN
		MANUAL_CHARACTER = "<img src='images/warning.gif' /> MANUAL - SELF COMPARING/CHECKING"
		'MANUAL_CHARACTER = "<img src='images/blank.png' class='icon manual' /> MANUAL - SELF COMPARING/CHECKING"
	ELSE
		MANUAL_CHARACTER = "(MANUAL)"
	END IF
END FUNCTION

FUNCTION NEWLINE_CHARACTER
	IF HTML_MODE = TRUE THEN
		NEWLINE_CHARACTER = "<br/>"
	ELSE
		NEWLINE_CHARACTER = vbcrlf
	END IF
END FUNCTION

FUNCTION UNDERCONSTRUCTION_CHARACTER
	IF HTML_MODE = TRUE THEN
		UNDERCONSTRUCTION_CHARACTER = "<img src='images/construct.gif' /> UNDER CONSTRUCTION"
		'UNDERCONSTRUCTION_CHARACTER = "<img src='images/blank.png' class='icon construction' /> UNDER CONSTRUCTION"
	ELSE
		UNDERCONSTRUCTION_CHARACTER = "(UNDER CONSTRUCTION)"
	END IF
END FUNCTION

PUBLIC SUB WRITE_TH(BYVAL FIELD1, BYVAL FIELD2, BYVAL FIELD3, BYVAL FIELD4, BYVAL FIELD5)
	IF HTML_MODE=FALSE THEN EXIT SUB
	
	IF FIELD1="" OR FIELD1="&nbsp;" THEN ADD_CLASS1 = " unsortable"
	IF FIELD2="" OR FIELD2="&nbsp;" THEN ADD_CLASS2 = " unsortable"
	IF FIELD3="" OR FIELD3="&nbsp;" THEN ADD_CLASS3 = " unsortable"
	IF FIELD4="" OR FIELD4="&nbsp;" THEN ADD_CLASS4 = " unsortable"
	IF FIELD5="" OR FIELD5="&nbsp;" THEN ADD_CLASS5 = " unsortable"
	
	wscript.stdout.write "<tr>"
	wscript.stdout.write "<th class='field1" & ADD_CLASS1 & "'>" & FIELD1 & "</th>"
	wscript.stdout.write "<th class='field2" & ADD_CLASS2 & "'>" & FIELD2 & "</th>"
	wscript.stdout.write "<th class='field3" & ADD_CLASS3 & "'>" & FIELD3 & "</th>"
	wscript.stdout.write "<th class='field4" & ADD_CLASS4 & "'>" & FIELD4 & "</th>"
	wscript.stdout.write "<th class='field5" & ADD_CLASS5 & "'>" & FIELD5 & "</th>"
	wscript.stdout.write "</tr>"
	wscript.stdout.write vbCRLF
END SUB
