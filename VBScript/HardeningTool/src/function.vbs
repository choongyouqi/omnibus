
	ANOTEXISTED = Array("NOT EXISTED")
	AREQ_TYPE_2 = Array("0:BUILTIN\Administrators:2032127", "0:NT AUTHORITY\SYSTEM:2032127")
	AREQ_TYPE_2C= Array("0:BUILTIN\Administrators:2032127", "0:NT AUTHORITY\SYSTEM:2032127", "0:\CREATOR OWNER:2032127")
	AREQ_TYPE_3 = Array("0:BUILTIN\Administrators:2032127", "0:NT AUTHORITY\SYSTEM:2032127", "0:NT AUTHORITY\Authenticated Users:1179817")
	AREQ_TYPE_4 = Array("0:BUILTIN\Administrators:2032127", "0:NT AUTHORITY\SYSTEM:2032127", "0:NT AUTHORITY\Authenticated Users:1179817", "0:\CREATOR OWNER:2032127")
	AREQ_TYPE_5 = Array("0:BUILTIN\Administrators:2032127", "0:NT AUTHORITY\SYSTEM:2032127", "0:NT AUTHORITY\Authenticated Users:1245631", "0:\CREATOR OWNER:2032127", "0:BUILTIN\Replicator:1245631")

Function CACLS_REPORT(ByVal AuditNo, ByVal ArrRequired, ByVal Filename)

	strTEMPValue = ""
	arrADDITIONAL = ""
	arrNOTFULFILLED = ""
	strCompare = ""
	
	
	strCurrentValue = getFilePermissions(Filename, strTEMPValue)
	
	IF strTEMPValue = "NOT EXISTED" THEN
		WRITE_OUTPUT_5 AuditNo, Filename, getHTMLFilePerm(Join(ArrRequired, ";")), strCurrentValue, NA_CHARACTER()
		EXIT FUNCTION
	END IF
	
	arrNOTFULFILLED = arrayFilter_sensitive(ArrRequired, split(strTEMPValue, ";"), arrADDITIONAL)
	
	IF Len(Join(arrNOTFULFILLED, NEWLINE_CHARACTER()))=0 AND Len(Join(arrADDITIONAL, NEWLINE_CHARACTER()))=0 THEN
		strCompare = CONFIRM_CHARACTER()
	ELSEIF Len(Join(arrNOTFULFILLED, NEWLINE_CHARACTER()))=0 THEN
		
		
		IF Join(ArrRequired,";") = Join(AREQ_TYPE_2,";") AND Join(arrADDITIONAL, ";") = "0:\CREATOR OWNER:2032127" THEN
			strCompare = CONFIRM_CHARACTER() & " CREATOR OWNER: FULL"
		ELSE
			strCompare = WARNING_CHARACTER() & vbcrlf
			strCompare = strCompare & "<b>ADDITIONAL:</b>" & NEWLINE_CHARACTER()
			strCompare = strCompare & getHTMLFilePerm(Join(arrADDITIONAL, ";"))
		END IF
	
	ELSE
		
		IF Join(arrNOTFULFILLED, ";") = "0:\CREATOR OWNER:2032127" THEN
			strCompare = CONFIRM_CHARACTER() & " NO CREATOR OWNER"
		ELSE
			strCompare = ERROR_CHARACTER() & vbcrlf
			strCompare = strCompare & "<b>NOT FULFILLED:</b>" & NEWLINE_CHARACTER()
			strCompare = strCompare & getHTMLFilePerm(Join(arrNOTFULFILLED, ";"))
			
			strCompare = strCompare & NEWLINE_CHARACTER() & NEWLINE_CHARACTER()
			strCompare = strCompare & WARNING_CHARACTER() & vbcrlf
			strCompare = strCompare & "<b>ADDITIONAL:</b>" & NEWLINE_CHARACTER()
			strCompare = strCompare & getHTMLFilePerm(Join(arrADDITIONAL, ";"))
		END IF
	END IF
	
	WRITE_OUTPUT_5 AuditNo, Filename, getHTMLFilePerm(Join(ArrRequired, ";")), strCurrentValue, strCompare
	'WRITE_OUTPUT_5 AuditNo, Filename, getHTMLFilePerm(Join(ArrRequired, ";")), strCurrentValue, strCompare & "<br/>" & Join(ArrRequired, ";") & "<br/>" & strTEMPValue
	'WRITE_OUTPUT_5 AuditNo, Filename, "1@", strCurrentValue, strCompare

End Function


Function CACLS_REPORT_SILENCE(ByVal AuditNo, ByVal ArrRequired, ByVal Filename)

	strTEMPValue = ""
	arrADDITIONAL = ""
	arrNOTFULFILLED = ""
	strCompare = ""
	
	strCurrentValue = getFilePermissions(Filename, strTEMPValue)
	
	IF strTEMPValue = "NOT EXISTED" THEN CACLS_REPORT_SILENCE = FALSE : EXIT FUNCTION
	
	arrNOTFULFILLED = arrayFilter_sensitive(ArrRequired, split(strTEMPValue, ";"), arrADDITIONAL)
	
	IF Len(Join(arrNOTFULFILLED, NEWLINE_CHARACTER()))=0 AND Len(Join(arrADDITIONAL, NEWLINE_CHARACTER()))=0 THEN
		CACLS_REPORT_SILENCE = TRUE
	ELSE
		CACLS_REPORT_SILENCE = FALSE
	END IF

End Function

Function arrayFilter_sensitive(ByVal Haystack, ByVal Needle, ByRef NOT_IN_LIST)
	ORIGINAL_HAYSTACK = Haystack
	STR_NOT_IN_LIST = ""
	NOT_IN_LIST = Array()
	
	'wscript.stderr.write vbcrlf 
	'wscript.stderr.write "UBOUND HAYSTACK   : " & ubound(Haystack) & vbcrlf
	'wscript.stderr.write "UBOUND NEEDLE     : " & ubound(Needle) & vbcrlf

	FOR x=0 to UBOUND(Needle)
	'wscript.stderr.write "PROCESSING NEEDLE : " & Needle(x) & vbcrlf
		IF NOT Needle(x) = "" THEN
	'wscript.stderr.write "NEEDLE FOUND. MATCHED NUMBER OF HAYSTACK: " & UBOUND(Filter(Haystack,Needle(x),true)) & vbcrlf 'filtered matched requirement
			IF UBOUND(Filter(Haystack,Needle(x),true)) = -1 THEN 'IF NOT MATCH THE REMAINING REQUIREMENT
				IF UBOUND(Filter(ORIGINAL_HAYSTACK,Needle(x),true)) = -1 THEN 'AND NOT MATCH THE ORIGINAL REQUIREMENT TOO (TO FILTER OUT DUPLICATION)
					STR_NOT_IN_LIST = STR_NOT_IN_LIST & Needle(x) & ";" 'NOT FOUND FROM THE LIST
					'wscript.stderr.write "NOT IN LIST TRIGGERED" & vbcrlf
				END IF
			END IF
			
			Haystack = Filter(Haystack,Needle(x),false) 'requirement left over
			'wscript.stderr.write JOIN(Haystack, " | ") & vbcrlf
			'wscript.stderr.write "FINAL UB(HAYSTACK): " & ubound(Haystack) & vbcrlf
	
		END IF
	'wscript.stderr.write vbcrlf
	NEXT
	arrayFilter_sensitive = Haystack
	IF NOT STR_NOT_IN_LIST = "" THEN NOT_IN_LIST = Split(Left(STR_NOT_IN_LIST,LEN(STR_NOT_IN_LIST)-1), ";")
End Function


Function arrayFilter(ByVal Haystack, ByVal Needle, ByRef NOT_IN_LIST)
	STR_NOT_IN_LIST = ""
	NOT_IN_LIST = Array()
	
	FOR x=0 to UBOUND(Needle)
	'wscript.stdout.write Needle(x)
		IF NOT Needle(x) = "" THEN
	'wscript.stdout.write vbtab & UBOUND(Filter(Haystack,Needle(x),true))
			TEMP_ARR = Filter(Haystack,Needle(x),true)
			IF UBOUND(TEMP_ARR) = -1 THEN STR_NOT_IN_LIST = STR_NOT_IN_LIST & Needle(x) & ";" 'NOT FOUND FROM THE LIST
			
			
			Haystack = Filter(Haystack,Needle(x),false)
			IF UBOUND(TEMP_ARR)>0 THEN
			
				FOR y=0 to UBOUND(TEMP_ARR)
					IF NOT TEMP_ARR(y)=Needle(x) THEN
						REDIM Preserve Haystack(UBOUND(Haystack)+1) 
						'wscript.stdout.write y
						 Haystack(UBOUND(Haystack)) = TEMP_ARR(y)
					END IF
				NEXT
			
			END IF
			
			'Haystack = JOIN(Haystack, " | ") & " | "
			
			'wscript.stderr.write "=====================================" & ubound(Haystack) & vbcrlf
		END IF
	'wscript.stdout.write vbcrlf
	NEXT
	arrayFilter = Haystack
	IF NOT STR_NOT_IN_LIST = "" THEN NOT_IN_LIST = Split(Left(STR_NOT_IN_LIST,LEN(STR_NOT_IN_LIST)-1), ";")
End Function

Function UnicodeToAnsi(inFile, outFile)
	' FileSystemObject.CreateTextFile and FileSystemObject.OpenTextFile
	Const OpenAsASCII   = 0 
	Const OpenAsUnicode = -1

	' FileSystemObject.CreateTextFile
	Const OverwriteIfExist = -1
	Const FailIfExist      = 0

	' FileSystemObject.OpenTextFile
	Const OpenAsDefault    = -2
	Const CreateIfNotExist = -1
	Const FailIfNotExist   = 0
	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8


	Set FileSys = CreateObject("Scripting.FileSystemObject")
	Set inStream  = FileSys.OpenTextFile(inFile, ForReading, FailIfNotExist, OpenAsDefault)
	Set outStream = FileSys.CreateTextFile(outFile, OverwriteIfExist, OpenAsASCII)
	Do
		inLine = inStream.ReadLine
		outStream.WriteLine inLine
	Loop Until inStream.AtEndOfStream
	inStream.Close
	outStream.Close
End Function


Function ReinterpretSignedAsUnsigned(ByVal x)
	If x < 0 Then x = x + 2^32
	ReinterpretSignedAsUnsigned = x
End Function

Function UnsignedDecimalStringToHex(ByVal x)
	x = CDbl(x)
	If x > 2^31 - 1 Then x = x - 2^32
	UnsignedDecimalStringToHex = Hex(x)
End Function

Function ISODate(datenow)
	'strDate = DatePart("yyyy-mm-dd hh:ii:ss AA", now())
	'datenow = now()
	strDate = 	DatePart("yyyy", datenow) & "-" & _
				Padding(DatePart("m", datenow), 2) & "-" & _
				Padding(DatePart("d", datenow), 2) & " " & _
				Padding(DatePart("h", datenow), 2) & ":" & _
				Padding(DatePart("n", datenow), 2) & ":" & _
				Padding(DatePart("s", datenow), 2)
	'WScript.Echo strDate
	ISODate = strDate
End Function

Function Exec(cmd)
	SET shell = WScript.CreateObject("WScript.Shell")
	'shell.run(cmd,,true)
	
	WriteLog(cmd)
		SET proc = shell.exec(cmd)
		DO WHILE proc.Status = 0: WScript.Sleep 100: Loop
	WriteLog("Status:" & proc.Status  & " ExitCode:"  & proc.ExitCode)
	IF proc.ExitCode <> 0 THEN WriteLog("STDERR:" & proc.StdErr.ReadAll)
	
	'proc.Status <> 0
	'proc.ExitCode <> 0 THEN
	'proc.StdErr.ReadAll
	'proc.StdOut.ReadAll
End Function

Function WriteLog(content)
	'If objFSO.FileExists(logfilename) Then objFSO.DeleteFile(logfilename)
	logfilename = "exec.log"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set f = objFSO.OpenTextFile(logfilename, 8, true)
	f.WriteLine(ISODate(now()) & " " & content)
	f.Close()
	Set f = Nothing
End Function

Function Padding(n, totalDigits) 
        Padding= Right(String(totalDigits,"0") & n, totalDigits) 
End Function 

Function BPadding(n, totalDigits) 
        BPadding= Right(String(totalDigits," ") & n, totalDigits) 
End Function 

Function FPadding(n, totalDigits) 
        FPadding= Left(n & String(totalDigits," "), totalDigits) 
End Function 

Function CFPadding(n, totalDigits) 
        FPadding= Left(n & String(totalDigits,"="), totalDigits) 
End Function 

Function ReadIni( myFilePath, mySection, myKey )
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
End Function

Sub WriteIni( myFilePath, mySection, myKey, myValue )
    ' This subroutine writes a value to an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be written
    ' myValue     [string]  the value to be written (myKey will be
    '                       deleted if myValue is <DELETE_THIS_VALUE>)
    '
    ' Returns:
    ' N/A
    '
    ' CAVEAT:     WriteIni function needs ReadIni function to run
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre, Johan Pol and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim blnInSection, blnKeyExists, blnSectionExists, blnWritten
    Dim intEqualPos
    Dim objFSO, objNewIni, objOrgIni, wshShell
    Dim strFilePath, strFolderPath, strKey, strLeftString
    Dim strLine, strSection, strTempDir, strTempFile, strValue

    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )
    strValue    = Trim( myValue )

    Set objFSO   = CreateObject( "Scripting.FileSystemObject" )
    Set wshShell = CreateObject( "WScript.Shell" )

    strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )

    Set objOrgIni = objFSO.OpenTextFile( strFilePath, ForReading, True )
    Set objNewIni = objFSO.CreateTextFile( strTempFile, False, False )

    blnInSection     = False
    blnSectionExists = False
    ' Check if the specified key already exists
    blnKeyExists     = ( ReadIni( strFilePath, strSection, strKey ) <> "" )
    blnWritten       = False

    ' Check if path to INI file exists, quit if not
    strFolderPath = Mid( strFilePath, 1, InStrRev( strFilePath, "\" ) )
    If Not objFSO.FolderExists ( strFolderPath ) Then
        WScript.Echo "Error: WriteIni failed, folder path (" _
                   & strFolderPath & ") to ini file " _
                   & strFilePath & " not found!"
        Set objOrgIni = Nothing
        Set objNewIni = Nothing
        Set objFSO    = Nothing
        WScript.Quit 1
    End If

    While objOrgIni.AtEndOfStream = False
        strLine = Trim( objOrgIni.ReadLine )
        If blnWritten = False Then
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                blnSectionExists = True
                blnInSection = True
            ElseIf InStr( strLine, "[" ) = 1 Then
                blnInSection = False
            End If
        End If

        If blnInSection Then
            If blnKeyExists Then
                intEqualPos = InStr( 1, strLine, "=", vbTextCompare )
                If intEqualPos > 0 Then
                    strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                    If LCase( strLeftString ) = LCase( strKey ) Then
                        ' Only write the key if the value isn't empty
                        ' Modification by Johan Pol
                        If strValue <> "<DELETE_THIS_VALUE>" Then
                            objNewIni.WriteLine strKey & "=" & strValue
                        End If
                        blnWritten   = True
                        blnInSection = False
                    End If
                End If
                If Not blnWritten Then
                    objNewIni.WriteLine strLine
                End If
            Else
                objNewIni.WriteLine strLine
                    ' Only write the key if the value isn't empty
                    ' Modification by Johan Pol
                    If strValue <> "<DELETE_THIS_VALUE>" Then
                        objNewIni.WriteLine strKey & "=" & strValue
                    End If
                blnWritten   = True
                blnInSection = False
            End If
        Else
            objNewIni.WriteLine strLine
        End If
    Wend

    If blnSectionExists = False Then ' section doesn't exist
        objNewIni.WriteLine
        objNewIni.WriteLine "[" & strSection & "]"
            ' Only write the key if the value isn't empty
            ' Modification by Johan Pol
            If strValue <> "<DELETE_THIS_VALUE>" Then
                objNewIni.WriteLine strKey & "=" & strValue
            End If
    End If

    objOrgIni.Close
    objNewIni.Close

    ' Delete old INI file
    objFSO.DeleteFile strFilePath, True
    ' Rename new INI file
    objFSO.MoveFile strTempFile, strFilePath

    Set objOrgIni = Nothing
    Set objNewIni = Nothing
    Set objFSO    = Nothing
    Set wshShell  = Nothing
End Sub















SID_ID = Array( _
"*S-1-0", _
"*S-1-0-0", _
"*S-1-1", _
"*S-1-1-0", _
"*S-1-2", _
"*S-1-3", _
"*S-1-3-0", _
"*S-1-3-1", _
"*S-1-3-2", _
"*S-1-3-3", _
"*S-1-4", _
"*S-1-5", _
"*S-1-5-1", _
"*S-1-5-2", _
"*S-1-5-3", _
"*S-1-5-4", _
"*S-1-5-5-X-Y", _
"*S-1-5-6", _
"*S-1-5-7", _
"*S-1-5-8", _
"*S-1-5-9", _
"*S-1-5-10", _
"*S-1-5-11", _
"*S-1-5-12", _
"*S-1-5-13", _
"*S-1-5-18", _
"*S-1-5-19", _
"*S-1-5-20", _
"*S-1-5-32-544", _
"*S-1-5-32-545", _
"*S-1-5-32-546", _
"*S-1-5-32-547", _
"*S-1-5-32-548", _
"*S-1-5-32-549", _
"*S-1-5-32-550", _
"*S-1-5-32-551", _
"*S-1-5-32-552", _
"*S-1-5-32-554", _
"*S-1-5-32-555", _
"*S-1-5-32-556", _
"*S-1-5-32-557", _
"*S-1-5-32-558", _
"*S-1-5-32-559", _
"*S-1-5-32-560", _
"*S-1-5-32-561", _
"*S-1-5-32-562", _
"*S-1-5-32-568", _
"*S-1-5-32-569", _
"*S-1-5-32-573", _
"*S-1-5-64-10", _
"*S-1-5-64-14", _
"*S-1-5-64-21", _
"*S-1-5-64-1000", _
"*S-1-6", _
"*S-1-7", _
"*S-1-8", _
"*S-1-9")


SID_NAME = Array( _
"Null Authority", _
"Nobody", _
"World Authority", _
"Everyone", _
"Local Authority", _
"Creator Authority", _
"Creator Owner", _
"Creator Group", _
"Creator Owner Server", _
"Creator Group Server", _
"Non-unique Authority", _
"NT Authority", _
"Dialup", _
"Network", _
"Batch", _
"Interactive", _
"Logon Session", _
"Service", _
"Anonymous", _
"Proxy", _
"Enterprise Domain Controllers", _
"Principal Self", _
"Authenticated Users", _
"Restricted Code", _
"Terminal Server Users", _
"Local System", _
"NT Authority (Local Service)", _
"NT Authority (Network Service)", _
"Administrators", _
"Users", _
"Guests", _
"Power Users", _
"Account Operators", _
"Server Operators", _
"Print Operators", _
"Backup Operators", _
"Replicators", _
"BUILTIN\Pre-Windows 2000 Compatible Access", _
"BUILTIN\Remote Desktop Users", _
"BUILTIN\Network Configuration Operators", _
"BUILTIN\Incoming Forest Trust Builders", _
"BUILTIN\Performance Monitor Users", _
"BUILTIN\Performance Log Users", _
"BUILTIN\Windows Authorization Access Group", _
"BUILTIN\Terminal Server License Servers", _
"BUILTIN\Distributed COM User", _
"BUILTIN\IIS_IUSRS", _
"BUILTIN\Cryptograhic Operators", _
"BUILTIN\Event Log Readers", _
"NTLM Authentication", _
"SChannel Authentication", _
"Digest Authentication", _
"Other Organization", _
"Site Server Authority An identifier authority.", _
"Internet Site Authority An identifier authority.", _
"Exchange Authority An identifier authority.", _
"Resource Manager Authority An identifier")


Function ReplaceSID(ByVal RawText)
	RawText= RawText & ","
	For replacesid_index = 0 to ubound(SID_ID)
		'wscript.stdout.write vbtab & Replace(RawText, SID_ID(x) & ",", SID_NAME(x) & ",") & vbcrlf
		RawText = Replace(RawText, SID_ID(replacesid_index) & ",", SID_NAME(replacesid_index) & ",")
	NEXT
	ReplaceSID = Left(RawText, Len(RawText)-1)
End Function


Function getFilePermissions(strFileName, ByRef inline_output)
	inline_output = ""
	Const SE_DACL_PRESENT = &h4
	Const ACCESS_ALLOWED_ACE_TYPE = &h0
	Const ACCESS_DENIED_ACE_TYPE  = &h1

	Const FILE_ALL_ACCESS       = &h1f01ff
	Const FILE_APPEND_DATA      = &h000004
	Const FILE_DELETE           = &h010000
	Const FILE_DELETE_CHILD     = &h000040
	Const FILE_EXECUTE          = &h000020
	Const FILE_READ_ATTRIBUTES  = &h000080
	Const FILE_READ_CONTROL     = &h020000
	Const FILE_READ_DATA        = &h000001
	Const FILE_READ_EA          = &h000008
	Const FILE_WRITE_ATTRIBUTES = &h000100
	Const FILE_WRITE_DAC        = &h040000
	Const FILE_WRITE_DATA       = &h000002
	Const FILE_WRITE_EA         = &h000010
	Const FILE_WRITE_OWNER      = &h080000

	Dim objWMIService, objAE,  intControlFlags, intRetVal, objFileSecuritySettings, strTemp, objSD, arrACEs
	Dim objFile, strSpecialPerms

	'getFilePermissions = strFileName & ":" '& vbCRLF '& vbTab

	Set objWMIService = GetObject("winmgmts:")

	Set objFileSecuritySettings = _
		objWMIService.ExecQuery("SELECT * FROM Win32_LogicalFileSecuritySetting WHERE Path='" & Replace(strFileName,"\","\\") & "'")

	For Each objFile in objFileSecuritySettings
		intRetVal = objFile.GetSecurityDescriptor(objSD)
		intControlFlags = objSD.ControlFlags

		If intControlFlags And SE_DACL_PRESENT Then
			arrACEs = objSD.DACL
	
			For Each objAE in arrACEs
				'STR_ACCESSMASK = getAccessMask(objAE.AccessMask)
				'IF objAE.AceType = ACCESS_ALLOWED_ACE_TYPE THEN additional_class = "font_black" ELSE additional_class = "font_red"
				'getFilePermissions = getFilePermissions & "<div class='clearboth " & additional_class & "'><div class='domainuser'>" & objAE.Trustee.Domain & "\" & objAE.Trustee.Name & ":</div><div class='accessmask'>" & STR_ACCESSMASK & "</div></div>"
				inline_output = inline_output & objAE.AceType & ":" & objAE.Trustee.Domain & "\" & objAE.Trustee.Name & ":" & objAE.AccessMask & ";"
				
			Next
			'getFilePermissions = getFilePermissions & vbcrlf
			'inline_output = inline_output & ";"
		Else
			getFilePermissions = getFilePermissions & "No DACL present in " & "security descriptor" & vbCRLF
		End If
	Next

		getFilePermissions = getFilePermissions & getHTMLFilePerm(inline_output)

	If(Len(getFilePermissions) = 0) Then
		getFilePermissions = getFilePermissions & "<div class='domainuser' style='color: blue; font-weight: 900;'>NOT EXISTED</div>"
		inline_output = "NOT EXISTED"
	END IF

	Set objAE = Nothing
	set arrACEs = Nothing
	Set objWMIService = Nothing
	Set objFileSecuritySettings = Nothing
	Set objFile = Nothing
End Function



Function getFilePermissionsArray(strFileName)
	inline_output = ""
	Const SE_DACL_PRESENT = &h4
	Const ACCESS_ALLOWED_ACE_TYPE = &h0
	Const ACCESS_DENIED_ACE_TYPE  = &h1

	Const FILE_ALL_ACCESS       = &h1f01ff
	Const FILE_APPEND_DATA      = &h000004
	Const FILE_DELETE           = &h010000
	Const FILE_DELETE_CHILD     = &h000040
	Const FILE_EXECUTE          = &h000020
	Const FILE_READ_ATTRIBUTES  = &h000080
	Const FILE_READ_CONTROL     = &h020000
	Const FILE_READ_DATA        = &h000001
	Const FILE_READ_EA          = &h000008
	Const FILE_WRITE_ATTRIBUTES = &h000100
	Const FILE_WRITE_DAC        = &h040000
	Const FILE_WRITE_DATA       = &h000002
	Const FILE_WRITE_EA         = &h000010
	Const FILE_WRITE_OWNER      = &h080000

	Dim objWMIService, objAE,  intControlFlags, intRetVal, objFileSecuritySettings, strTemp, objSD, arrACEs
	Dim objFile, strSpecialPerms

	'getFilePermissions = strFileName & ":" '& vbCRLF '& vbTab

	Set objWMIService = GetObject("winmgmts:")

	Set objFileSecuritySettings = _
		objWMIService.ExecQuery("SELECT * FROM Win32_LogicalFileSecuritySetting WHERE Path='" & Replace(strFileName,"\","\\") & "'")

	For Each objFile in objFileSecuritySettings
		intRetVal = objFile.GetSecurityDescriptor(objSD)
		intControlFlags = objSD.ControlFlags

		If intControlFlags And SE_DACL_PRESENT Then
			arrACEs = objSD.DACL
			For Each objAE in arrACEs
				STR_ACCESSMASK = getAccessMask(objAE.AccessMask)
				inline_output = inline_output & objAE.AceType & ":" & objAE.Trustee.Domain & "\" & objAE.Trustee.Name & ":" & objAE.AccessMask & ";"
			Next
		End If
	Next

	'If(Len(getFilePermissionsArray) = 0) Then
	'	inline_output = "NOT EXISTED"
	'END IF

	Set objAE = Nothing
	set arrACEs = Nothing
	Set objWMIService = Nothing
	Set objFileSecuritySettings = Nothing
	Set objFile = Nothing
	
	getFilePermissionsArray = split(inline_output, ";")
End Function


Function getAccessMask(accessmask_code)

	Const SEPERATOR_CHARACTER     = "<br/>" '", "
	
	Const FILE_ALL_ACCESS         = &h1f01ff
	Const FILE_APPEND_DATA        = &h000004
	Const FILE_DELETE             = &h010000
	Const FILE_DELETE_CHILD       = &h000040
	Const FILE_EXECUTE            = &h000020
	Const FILE_READ_ATTRIBUTES    = &h000080
	Const FILE_READ_CONTROL       = &h020000
	Const FILE_READ_DATA          = &h000001
	Const FILE_READ_EA            = &h000008
	Const FILE_WRITE_ATTRIBUTES   = &h000100
	Const FILE_WRITE_DAC          = &h040000
	Const FILE_WRITE_DATA         = &h000002
	Const FILE_WRITE_EA           = &h000010
	Const FILE_WRITE_OWNER        = &h080000
	
	Const FILE_GENERIC_READALL    = &h60000000

	
	SELECT CASE accessmask_code
	CASE 268435456: getAccessMask = "<font color='red'>SPECIAL:</font> FC"
	
	CASE 2032127: getAccessMask = "FULL CONTROL"
	CASE 1245631: getAccessMask = "MODIFY"
	CASE 1179817: getAccessMask = "READ & EXEC"
	CASE 1179785: getAccessMask = "READ"
	CASE 1048854: getAccessMask = "WRITE"
	CASE 1180063: getAccessMask = "READ & WRITE"
	CASE 1180095: getAccessMask = "READ, WRITE & EXEC"
	CASE ELSE
		'getAccessMask = "SPECIAL:" & accessmask_code
		getAccessMask = "<font color='red'>SPECIAL:</font> "
		If accessmask_code And FILE_GENERIC_READALL Then getAccessMask = "RD_ALL" & SEPERATOR_CHARACTER
		
		If accessmask_code And FILE_APPEND_DATA Then getAccessMask = getAccessMask & "Append Data" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_DELETE Then getAccessMask = getAccessMask & "Delete" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_EXECUTE Then getAccessMask = getAccessMask & "Execute File" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_READ_ATTRIBUTES Then getAccessMask = getAccessMask & "Read Attributes" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_READ_CONTROL Then getAccessMask = getAccessMask & "Read Permissions" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_READ_DATA Then getAccessMask = getAccessMask & "Read Data" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_READ_EA Then getAccessMask = getAccessMask & "Read Attributes" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_WRITE_ATTRIBUTES Then getAccessMask = getAccessMask & "Write Attributes" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_WRITE_DAC Then getAccessMask = getAccessMask & "Change " & "Permissions" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_WRITE_DATA Then getAccessMask = getAccessMask & "Write Data" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_WRITE_EA Then getAccessMask = getAccessMask & "Write Extended " & "Attributes" & SEPERATOR_CHARACTER
		If accessmask_code And FILE_WRITE_OWNER Then getAccessMask = getAccessMask & "Take Ownership" & SEPERATOR_CHARACTER
		getAccessMask = LEFT(getAccessMask, LEN(getAccessMask)-2)
		'IF getAccessMask = "SPECIAL" THEN getAccessMask = "SPECIAL: " & accessmask_code
	END SELECT

End Function


Function getHTMLFilePerm(raw_string)
	IF raw_string = "NOT EXISTED" THEN getHTMLFilePerm = "<div class='domainuser' style='color: blue; font-weight: 900;'>NOT EXISTED</div>": EXIT FUNCTION
	Const SE_DACL_PRESENT = &h4
	Const ACCESS_ALLOWED_ACE_TYPE = &h0
	Const ACCESS_DENIED_ACE_TYPE  = &h1
	
	raw_array = split(raw_string, ";")
	FOR x = 0 TO ubound(raw_array)
		IF NOT raw_array(x) = "" THEN
			
			sub_array  = split(raw_array(x), ":")
			acetype    = sub_array(0)
			domainuser = sub_array(1)
			accessmask = sub_array(2)
			
			'wscript.stdout.write "acetype:" & vbtab & acetype & vbcrlf
			'wscript.stdout.write "domainuser:" & vbtab & domainuser & vbcrlf
			'wscript.stdout.write "accessmask:" & vbtab & getAccessMask(accessmask) & vbcrlf
			'wscript.stdout.write vbcrlf
			
			IF CINT(acetype) = ACCESS_ALLOWED_ACE_TYPE THEN additional_class = "font_black" ELSE additional_class = "font_red"
			getHTMLFilePerm = getHTMLFilePerm & "<div class='clearboth " & additional_class & "'>"
			getHTMLFilePerm = getHTMLFilePerm & "<div class='domainuser'>" & domainuser & ":</div>"
			getHTMLFilePerm = getHTMLFilePerm & "<div class='accessmask'>" & getAccessMask(accessmask) & "</div>"
			getHTMLFilePerm = getHTMLFilePerm & "</div>"
		END IF
	NEXT

End Function
