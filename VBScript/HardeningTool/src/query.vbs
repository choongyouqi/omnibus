'ON ERROR RESUME NEXT


'================================================================================
'  DEFINE CONSTANTS (taken from WinReg.h)
'================================================================================
	Const HKEY_CLASSES_ROOT   = &H80000000
	Const HKEY_CURRENT_USER   = &H80000001
	Const HKEY_LOCAL_MACHINE  = &H80000002
	Const HKEY_USERS          = &H80000003
	Const REG_SZ        = 1
	Const REG_EXPAND_SZ = 2
	Const REG_BINARY    = 3
	Const REG_DWORD     = 4
	Const REG_MULTI_SZ  = 7
	LEN_NAME = 20

'================================================================================
'  IN BETWEEN CONSTANT
'================================================================================
	'oReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath
	'oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,arrKeyNames,strValue
	'oReg.DeleteValue HKEY_LOCAL_MACHINE,strKeyPath,strDWORDValueName
	'oReg.DeleteKey HKEY_LOCAL_MACHINE, strKeyPath

	'hDefKey = HKEY_LOCAL_MACHINE
	'strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Run"
	
	'hDefKey = HKEY_CURRENT_USER
	'strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Run"
	

'================================================================================
'  GET ARGUMENT
'================================================================================
	SET args = WScript.Arguments
	IF NOT args.Count = 0 THEN
		LEN_NAME = args.Item(0)
	END IF

'================================================================================
'  INCLUDE FUNCTIONS
'================================================================================
	SET objFSO = CreateObject("Scripting.FileSystemObject")
	SET funcinc = objFSO.OpenTextFile("function.vbs", 1)
	Execute funcinc.ReadAll
	funcinc.Close

	SET formatinc = objFSO.OpenTextFile("format.vbs", 1)
	Execute formatinc.ReadAll
	formatinc.Close

'================================================================================
'  LOAD CONFIGURATIONS
'================================================================================

	Dim HTML_MODE
	SET objShell = CreateObject("Wscript.Shell")
	DIRECTORY_CURRENT = objShell.CurrentDirectory
	CONFIG_FILENAME = DIRECTORY_CURRENT & "\config.ini"
	HTML_MODE  = CBOOL(ReadInI(CONFIG_FILENAME, "SERVERSETTING", "HTML_MODE"))
	SKIP_10_6  = CBOOL(ReadInI(CONFIG_FILENAME, "SERVERSETTING", "SKIP_10_6"))
	SKIP_ERROR = CBOOL(ReadInI(CONFIG_FILENAME, "SERVERSETTING", "SKIP_ERROR"))
	JAVASCRIPT = CBOOL(ReadInI(CONFIG_FILENAME, "SERVERSETTING", "JAVASCRIPT"))
	SILENCE    = CBOOL(ReadInI(CONFIG_FILENAME, "SERVERSETTING", "SILENCE"))
	
	Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	hDefKey = HKEY_LOCAL_MACHINE



'================================================================================
'  DOCUMENT PREFIX/HEADER
'================================================================================

IF SKIP_ERROR = TRUE THEN ON ERROR RESUME NEXT

IF HTML_MODE = TRUE THEN
	wscript.stdout.write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbcrlf
	wscript.stdout.write "<html><head>" & vbcrlf
	wscript.stdout.write "<title>" & CreateObject("WScript.Network").Computername & " - SERVER REVIEWING REPORT</title>" & vbcrlf
	wscript.stdout.write "<link rel=""stylesheet"" href=""images/style.css"" type=""text/css"" media=""screen"" />" & vbcrlf
	wscript.stdout.write "<link rel=""stylesheet"" href=""images/style.css"" type=""text/css"" media=""print"" />" & vbcrlf
	IF JAVASCRIPT = TRUE THEN wscript.stdout.write "<script type=""text/javascript"" src=""images/sortable.js""></script>" & vbcrlf
	wscript.stdout.write "</head><body><div id=""content"">" & vbcrlf
	wscript.stdout.write "<div id=header><img src=""images/osa_user_audit.png"" /> WINDOWS SERVER REVIEWING REPORT</div>" & vbcrlf
END IF





'================================================================================
'  SECTION 10 (TESTING ENVIRONMENT)
'================================================================================
'WRITE_OUTPUT_SECTION_START "SECTION 10"
'WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"


	'Set objEnv = objShell.Environment("Process")
	
	'AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\Documents and Settings"
	'CACLS_REPORT "10.7", AREQ_TYPE_4, AUDIT_FILENAME
	
	
	'AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\Documents and Settings\Administrator"
	'CACLS_REPORT "10.8", AREQ_TYPE_2, AUDIT_FILENAME
	
	'AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\config"
	'CACLS_REPORT "10.13", AREQ_TYPE_2, AUDIT_FILENAME
	
	'AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\io.sys"
	'CACLS_REPORT "10.26", AREQ_TYPE_3, AUDIT_FILENAME
	
	'wscript.quit


'WRITE_OUTPUT_SECTION_END "SECTION 10"




	
'================================================================================
'  GET HOSTNAME AND IP ADDRESS
'================================================================================
WRITE_OUTPUT_SECTION_START_COL "SERVER INFORMATION", "", 5
WRITE_TH "", "Machine Information", "Current Value", "&nbsp;", "&nbsp;"

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objWSHNetwork = CreateObject("WScript.Network")
	Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
	Set colIPResults = objWMI.ExecQuery("SELECT IPAddress, MacAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'True'")
	strMACAddress = ""
	For Each objNIC In colIPResults
		strMACAddress = strMACAddress & objNIC.MacAddress
		For Each strIPAddress in objNIC.IPAddress
			IF strAddresses = "" Then strAddresses = strIPAddress Else strAddresses = strAddresses
		Next
	Next
	strHostname = objWSHNetwork.Computername
	
	WRITE_OUTPUT_5 "", "HOSTNAME", strHostname, "", ""
	WRITE_OUTPUT_5 "", "IP ADDRESS", strAddresses, "", ""
	WRITE_OUTPUT_5 "", "MAC ADDRESS", strMACAddress, "", ""
	WRITE_OUTPUT_5 "", "DATE TIME", ISODate(now()), "", ""
	
	Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	For Each objItem In colItems

		'WRITE_OUTPUT_5 "", "COMPUTER NAME", objItem.Name, "", ""
		'WRITE_OUTPUT_5 "", "NAME FORMAT", objItem.NameFormat, "", ""
		WRITE_OUTPUT_5 "", "DOMAIN", objItem.Domain, "", ""
		WRITE_OUTPUT_5 "", "PART OF DOMAIN", objItem.PartOfDomain, "", "" 'post-Windows 2000 only
		WRITE_OUTPUT_5 "", "WORKGROUP", objItem.Workgroup, "", "" 'post-Windows 2000 only

		Select Case objItem.DomainRole
			Case 0 strDomainRole = "Standalone Workstation"
			Case 1 strDomainRole = "Member Workstation"
			Case 2 strDomainRole = "Standalone Server"
			Case 3 strDomainRole = "Member Server"
			Case 4 strDomainRole = "Backup Domain Controller"
			Case 5 strDomainRole = "Primary Domain Controller"
		End Select
		
		WRITE_OUTPUT_5 "", "DOMAIN ROLE", strDomainRole, "", ""
		WRITE_OUTPUT_5 "", "ROLES", Join(objItem.Roles, ","), "", ""
		WRITE_OUTPUT_5 "", "NETWORK SERVER MODE ENABLED", objItem.NetworkServerModeEnabled, "", ""
	
	Next
	
	
	Set oShell = CreateObject( "WScript.Shell" )
	PROCESSOR_ARCHITECTURE=oShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
		WRITE_OUTPUT_5 "", "PROCESSOR ARCHITECTURE", PROCESSOR_ARCHITECTURE, "", ""
	
WRITE_OUTPUT_SECTION_END "SERVER INFORMATION"

'================================================================================
'  SECTION 1
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 1 - Setup & Installation", ""
WRITE_TH "No", "Task/Description", "Required Value", "Current Value", "Remark"

	Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("Select * from Win32_OperatingSystem",,48)


	'Set colQuickFixes = objWMI.ExecQuery("SELECT * FROM Win32_QuickFixEngineering")
	'For Each objQuickFix in colQuickFixes
	'	If(StrComp(objQuickFix.HotFixID, "File 1") <> 0) Then
	'		listHotFixes = listHotFixes & "Hot Fix ID: " & objQuickFix.HotFixID & vbCRLF
	'	End If
	'Next
	'WRITE_OUTPUT_SECTION1 "1.0", "HOT FIXES:", listHotFixes
	
    'Get the OS version number (first two) and OS product type (server or desktop)
    For Each objItem in colItems
        OSVersion = Left(objItem.Version,3)
        ProductType = objItem.ProductType
        ServicePack = objItem.ServicePackMajorVersion & "." & objItem.ServicePackMinorVersion
    Next

	
	requireValue = "SP2.0"
	currentValue = "SP" & ServicePack
	IF currentValue=requireValue THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER() 
	WRITE_OUTPUT_5 "1.3", "SERVICE PACK", requireValue, currentValue, strCompare
	
	requireValue = "6.0.0.2406"
	IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Altiris\Altiris Agent", "Version", version) = 0 THEN
		currentValue = version
		strCompare = CONFIRM_CHARACTER()
	ELSEIF oReg.GetStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Altiris\Altiris Agent", "Version", version) = 0 THEN
		currentValue = version
		strCompare = CONFIRM_CHARACTER()
	ELSE
		currentValue = "NOT FOUND"
		strCompare = ERROR_CHARACTER()
	END IF
	
	WRITE_OUTPUT_5 "", "ALTIRIS", requireValue, currentValue, strCompare
	
	
	requireValue = "3151030"
	IF oReg.GetDwordValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Symantec\Symantec Endpoint Protection\AV", "ProductVersion", version) = 0 THEN
		currentValue = version
	ELSEIF oReg.GetDwordValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Symantec\Symantec Endpoint Protection\AV", "ProductVersion", version) = 0 THEN
		currentValue = version
		requireValue = "19011542"
	ELSE
		currentValue = "NOT FOUND"
	END IF
	
	IF cstr(currentValue)=requireValue THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER() 
	WRITE_OUTPUT_5 "", "SYMANTEC ENDPOINT PROTECTION", requireValue, currentValue, strCompare
	
	IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\McAfee\DesktopProtection", "szProductVer", version) = 0 THEN currentValue = version ELSE currentValue = "NOT FOUND"
	requireValue = "NOT FOUND"
	IF currentValue=requireValue THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER() 
	WRITE_OUTPUT_5 "", "McAfee", requireValue, currentValue, strCompare
	
    'Time to convert numbers into names
    Select Case OSVersion
		Case "7.0" : OSName = "Windows 7"
		Case "6.0" : OSName = "Windows Vista"
		Case "5.2" : OSName = "Windows 2003"
		Case "5.1" : OSName = "Windows XP"
		Case "5.0" : OSName = "Windows 2000"
		Case "4.0" : OSName = "Oh! Really! NT 4.0"
		Case Else : OSName = "Hi! Grandpa! Windows ME or older"
    End Select

	WRITE_OUTPUT_5 "1.4", "OPERATING SYSTEM", "", OSName, CONFIRM_CHARACTER()
	
	DriveType = Array("Unknown", "No Root Directory", "Removable Disk", "Local Disk", "Network Drive", "Compact Disc", "RAM Disk")
	value = ""
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk WHERE DriveType=3")
	For Each objDisk in colDisks
		
		IF HTML_MODE = TRUE THEN
		value = value & FPadding(objDisk.DeviceID, 5) & "\ (" & objDisk.FileSystem & ")<br/>"
		ELSE
		value = value & FPadding(objDisk.DeviceID, 5) & "\ (" & objDisk.FileSystem & "), " '& vbtab & DriveType(objDisk.DriveType) 
		END IF
		
	Next
	
	IF HTML_MODE = TRUE THEN value = Left(value, Len(value)-5) ELSE value = Left(value, Len(value)-2)
	WRITE_OUTPUT_5 "1.5", "FILE SYSTEM: ", "", value, CONFIRM_CHARACTER()
	
	
	
	
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'IF objFSO.FileExists("%ProgramFiles%\Outlook Express\msimn.exe") Then
	
	Set oShell = CreateObject( "WScript.Shell" )
	ProgramFiles=oShell.ExpandEnvironmentStrings("%ProgramFiles%")
	IF objFSO.FileExists(ProgramFiles & "\Outlook Express\msimn.exe") Then currentValue = "FOUND" ELSE currentValue = "NOT FOUND"
	requireValue = "NOT FOUND"
	IF currentValue=requireValue THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = WARNING_CHARACTER() & " No Uninstaller!"
	WRITE_OUTPUT_5 "1.6", "OUTLOOK EXPRESS", requireValue, currentValue, strCompare
	
	'http://support.microsoft.com/kb/q240794/
	strKeyPath = "SOFTWARE\Classes\Outlook.Application\CLSID"
	IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, "", currentValue) <> 0 Then
		currentValue = "NOT FOUND"
	ELSE
		strKeyPath = "SOFTWARE\Classes\CLSID\" & value & "\LocalServer32"
		IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, "", currentValue) <> 0 Then currentValue = "NOT FOUND" ELSE currentValue = "FOUND"
	END IF
	
	requireValue = "NOT FOUND"
	IF currentValue=requireValue THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER() & " SECURITY ALERT!"
	WRITE_OUTPUT_5 "", "OUTLOOK", requireValue, currentValue, strCompare
	
	
	
	
	
	
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colServiceList = objWMIService.ExecQuery("SELECT * FROM Win32_Service WHERE Name='wuauserv'")
	For Each objService in colServiceList
		strRequire = "Stopped | Disabled"
		strCurrent = objService.State & " | " & objService.StartMode
		IF strCurrent=strRequire THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
		WRITE_OUTPUT_5 "1.7", "AUTOMATIC UPDATE", strRequire, strCurrent, strCompare
	Next

WRITE_OUTPUT_SECTION_END "SECTION1"


'================================================================================
'  SECTION 2.2 THE FOLLOWING SERVICES MUST BE FLAGGED AS DISABLE
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 2.2 - Services", "<b>Disable</b> the following services"
WRITE_TH "No", "Task/Description", "Required Value", "Current Value", "Remark"

	SERV_CODE = Array("AppMgmt", "wuauserv", "BITS", "ClipSrv", "Dhcp", "TrkSvr", "TrkWks", "MSDTC", "ERSvc", "NtFrs", "helpsvc", "HTTPFilter", "HidServ", "ImapiService", "CiSvc", "IsmServ", "kdc", "LicenseService", "Messenger", "mnmsrvc", "NetDDE", "NetDDEdsdm", "WmdmPmSN", "Spooler", "RasAuto", "RasMan", "RDSessMgr", "RpcLocator", "RemoteAccess", "seclogon", "ShellHWDetection", "SCardSvr", "sacsvr", "TapiSrv", "TlntSvr", "Tssdis", "Themes", "UPS", "vds", "WebClient", "AudioSrv", "SharedAccess", "stisvc", "UMWdf", "WinHttpAutoProxySvc", "WZCSVC")
	SERV_DESC = Array("Application Management", "Automatic Updates", "Background Intelligent Transfer Service", "Clipbook", "DHCP Client", "Distributed Link Tracking Server", "Distributed Link Tracking Client", "Distributed Transaction Coordinator", "Error Reporting Services", "File Replication ", "Help and Support", "HTTP SSL", "Human Interface Device Access", "IMAPI CD-Burning Com Service", "Indexing Services", "Intersite Messaging", "Kerberos Key Distribution", "License Logging Service", "Messenger", "Netmeeting Remote Desktop Sharing", "Network DDE", "Network DDE DSDM", "Portable Media Serial Number Service", "Print Spooler", "Remote Access Auto Connection Manager", "Remote Access Connection Manager", "Remote Desktop Help Session Manager", "Remote Procedure Call (RPC) Locator", "Routing and Remote Access", "Secondary Logon", "Shell Hardware Detection", "Smart Card", "Special Admin Console Helper", "Telephony", "Telnet", "Terminal Services Session Directory", "Themes", "Uninterruptible Power Supply", "Virtual Disk Services", "WebClient", "Windows Audio", "Windows Firewall/Internet Connection Sharing", "Windows Image Acquisition", "Windows User Mode Driver Framework", "WinHTTP Web Proxy Auto Discover Services", "Wireless Configuration")
	SERV_STAT = Array("ERR 0" ,"ERR 1", "AUTO", "MANUAL", "DISABLED", "NOT EXISTED")

	FOR x = 0 TO ubound(SERV_CODE)
		'wscript.echo SERV_CODE(x) & vbtab & SERV_DESC(x) & vbcrlf
		'wscript.echo len(SERV_DESC(x))
		strKeyPath = "System\CurrentControlSet\Services\" & SERV_CODE(x)
		IF oReg.GetDWORDValue(hDefKey, strKeyPath, "Start", value) <> 0 Then value = 5
		
		IF value=4 THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
		IF value=5 THEN strCompare = NA_CHARACTER()
		
		IF (x+1)=12 THEN VALUE_HTTPSSL = value
		requireValue = "DISABLED"
		currentValue = SERV_STAT(value)
		WRITE_OUTPUT_5 "2.2." & (x+1), SERV_DESC(x), requireValue, currentValue, strCompare
	NEXT

WRITE_OUTPUT_SECTION_END "SECTION 2.2"


'================================================================================
'  SECTION 2.3 THE FOLLOWING SERVICES MUST BE FLAGGED AS MANUAL
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 2.3 - Services", "Ensure the following services are set to <b>Manual</b>"
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"

	SERV_CODE = Array("AeLookupSvc", "ALG", "EventSystem", "COMSysApp", "DcomLaunch", "Dfs", "dmserver", "dmadmin", "swprv", "Netman", "Nla", "xmlprov", "NtLmSsp", "SysmonLog", "NtmsSvc", "RSoPProv", "SNMPTRAP", "TermService", "VSS", "MSIServer", "Wmi", "WmiApSrv")
	SERV_DESC = Array("Application Experience Lookup Service", "Application Layer Gateway Services", "COM+ Event System", "COM+ System Application", "DCOM Server Process Launcher", "Distributed File System", "Logical Disk Manager", "Logical Disk Manager Administrative Service", "Microsoft Software Shadow Copy Provider", "Network Connection", "Network Location Awareness", "Network Provisioning Service", "NT LMD Security Support Provider", "Performance Logs And Alerts", "Removable Storage", "Resultant Set of Policy Provider", "SNMP Trap Service", "Terminal Services", "Volume Shadow Copy", "Windows Installer", "WMI Driver Extensions", "WMI Performance Adapter")
	
	FOR x = 0 TO ubound(SERV_CODE)
		'wscript.echo SERV_CODE(x) & vbtab & SERV_DESC(x) & vbcrlf
		'wscript.echo len(SERV_DESC(x))
		strKeyPath = "System\CurrentControlSet\Services\" & SERV_CODE(x)
		IF oReg.GetDWORDValue(hDefKey, strKeyPath, "Start", value) <> 0 THEN value = 5
		
		IF value=3 THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
		IF value=5 THEN strCompare = NA_CHARACTER()
		
		IF x=4 THEN 
			IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Altiris\Altiris Agent", "MachineGuid", value2) <> 0 THEN
				value2 = "NOT INSTALLED"
			ELSE
				strCompare = CONFIRM_CHARACTER() & " AUTO - ALTIRIS INSTALLED"
			END IF
		END IF

		requireValue = "MANUAL"
		currentValue = SERV_STAT(value)
		WRITE_OUTPUT_5 "2.3." & (x+1), SERV_DESC(x), requireValue, currentValue, strCompare
	NEXT

WRITE_OUTPUT_SECTION_END "SECTION 2.3"




'================================================================================
'  SECTION 3.1
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 3.1 - Password Policy", "Local Security Settings > Account Policies > Password Policy"
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"
  
	'SET objShell = CreateObject("Wscript.Shell")
	CurrentDir = objShell.CurrentDirectory
	SECPOL_FILENAME = CurrentDir & "\secpol.inf"

	CMD_RUN = "secedit.exe /export /cfg """ & SECPOL_FILENAME & """"
	objShell.run CMD_RUN, 0, true
	
	'wscript.stdout.write CurrentDir & "\secpol_ansi.inf" 
	UnicodeToAnsi SECPOL_FILENAME, CurrentDir & "\secpol_ansi.inf"
	IF objFSO.FileExists(SECPOL_FILENAME) THEN objFSO.DeleteFile SECPOL_FILENAME
	SECPOL_FILENAME = CurrentDir & "\secpol_ansi.inf"
	
	'wscript.quit
	'secedit.exe /export /cfg test2.inf
	'wscript.stdout.write CurrentDir & "\secpol.inf"
	'SET objFSO = CreateObject("Scripting.FileSystemObject")
	'SECPOL_FILENAME = "boot.ini"



	TASK31 = Array("PasswordHistorySize", "MaximumPasswordAge", "MinimumPasswordAge", "MinimumPasswordLength", "RequireLogonToChangePassword", "PasswordComplexity", "ClearTextPassword")
	TASK31_REQUIRED = Array("24", "45", "0", "8", "0", "0", "0")
	Dim TASK31_QUERY(9)
	FOR x = 0 TO ubound(TASK31)
		TASK31_QUERY(x) = ReadInI(SECPOL_FILENAME, "System Access", TASK31(x))
	NEXT
	
	
	
	TASK32 = Array("LockoutDuration", "LockoutBadCount", "ResetLockoutCount")
	TASK32_REQUIRED = Array("-1","3","10080")
	Dim TASK32_QUERY(9)
	FOR x = 0 TO ubound(TASK32)
		TASK32_QUERY(x) = ReadInI(SECPOL_FILENAME, "System Access", TASK32(x))
	NEXT

	'task_321 = ReadInI(SECPOL_FILENAME, "System Access", "LockoutDuration") '0
	'task_322 = ReadInI(SECPOL_FILENAME, "System Access", "LockoutBadCount") '3
	'task_323 = ReadInI(SECPOL_FILENAME, "System Access", "ResetLockoutCount") '10080
	
	task_41 = ReadInI(SECPOL_FILENAME, "System Access", "EnableAdminAccount") '1
	task_42 = ReadInI(SECPOL_FILENAME, "System Access", "EnableGuestAccount") '0
	task_44 = StrReverse(Mid(StrReverse(Mid(ReadInI(SECPOL_FILENAME, "System Access", "NewAdministratorName"),2)),2)) 'CUSTOMADMIN
	task_45 = StrReverse(Mid(StrReverse(Mid(ReadInI(SECPOL_FILENAME, "System Access", "NewGuestName"),2)),2)) 'CUSTOMGUEST
	task_443 = ReadInI(SECPOL_FILENAME, "System Access", "LSAAnonymousNameLookup") '0
	task_455 = ReadInI(SECPOL_FILENAME, "System Access", "ForceLogoffWhenHourExpire") '0
	
	EVENT_STATUS = Array("ERROR", "Success", "Failure", "Success, Failure")
	TASK5_REQUIRED = Array("3","3","2","3","2","1","3","2","1")
	TASK5 = Array("AuditAccountLogon", "AuditAccountManage", "AuditDSAccess", "AuditLogonEvents", "AuditObjectAccess", "AuditPolicyChange", "AuditPrivilegeUse", "AuditProcessTracking", "AuditSystemEvents")
	Dim TASK5_QUERY(9)
	FOR x = 0 TO ubound(TASK5)
		TASK5_QUERY(x) = ReadInI(SECPOL_FILENAME, "Event Audit", TASK5(x))
	NEXT

	TASK33_DESC = Array("Access this computer from the network", _
						"Act as part of the operating system", _
						"Add workstations to domains", _
						"Adjust memory quotas for a process", _
						"Allow log on locally", _
						"Allow log on through Terminal Services", _
						"Back up files and directories", _
						"Bypass traverse checking", _
						"Change the system time", _
						"Create a pagefile", _
						"Create a token object", _
						"Create global objects", _
						"Create permanent shared objects", _
						"Debug programs", _
						"Deny access to this computer from the network", _
						"Deny log on as a batch job", _
						"Profile single process", _
						"Remove computer from docking station", _
						"Restore files and directories", _
						"Shut down the system")
	TASK33 = Array("SeNetworkLogonRight", "SeTcbPrivilege", "SeMachineAccountPrivilege", "SeIncreaseQuotaPrivilege", "SeInteractiveLogonRight", "SeRemoteInteractiveLogonRight", "SeBackupPrivilege", "SeChangeNotifyPrivilege", "SeSystemTimePrivilege", "SeCreatePagefilePrivilege", "SeCreateTokenPrivilege", "SeCreateGlobalPrivilege", "SeCreatePermanentPrivilege", "SeDebugPrivilege", "SeDenyNetworkLogonRight", "SeDenyBatchLogonRight", "SeProfileSingleProcessPrivilege", "SeUndockPrivilege", "SeRestorePrivilege", "SeShutdownPrivilege")
	Dim TASK33_QUERY(20)
	FOR x = 0 TO ubound(TASK33)
		TASK33_QUERY(x) = ReadInI(SECPOL_FILENAME, "Privilege Rights", TASK33(x))
	NEXT



	'TASK08 = Array("SeNetworkLogonRight", "SeTcbPrivilege", "SeMachineAccountPrivilege", "SeIncreaseQuotaPrivilege", "SeInteractiveLogonRight", "SeRemoteInteractiveLogonRight", "SeBackupPrivilege", "SeChangeNotifyPrivilege", "SeSystemTimePrivilege", "SeCreatePagefilePrivilege", "SeCreateTokenPrivilege", "SeCreateGlobalPrivilege", "SeCreatePermanentPrivilege", "SeDebugPrivilege", "SeDenyNetworkLogonRight", "SeDenyBatchLogonRight", "SeProfileSingleProcessPrivilege", "SeUndockPrivilege", "SeRestorePrivilege", "SeShutdownPrivilege")
	TASK08 = Array("SeNetworkLogonRight", "SeMachineAccountPrivilege", "SeBackupPrivilege", "SeChangeNotifyPrivilege", "SeSystemtimePrivilege", "SeCreatePagefilePrivilege", "SeDebugPrivilege", "SeRemoteShutdownPrivilege", "SeAuditPrivilege", "SeIncreaseQuotaPrivilege", "SeIncreaseBasePriorityPrivilege", "SeLoadDriverPrivilege", "SeBatchLogonRight", "SeServiceLogonRight", "SeInteractiveLogonRight", "SeSecurityPrivilege", "SeSystemEnvironmentPrivilege", "SeProfileSingleProcessPrivilege", "SeSystemProfilePrivilege", "SeAssignPrimaryTokenPrivilege", "SeRestorePrivilege", "SeShutdownPrivilege", "SeTakeOwnershipPrivilege", "SeDenyNetworkLogonRight", "SeDenyBatchLogonRight", "SeDenyInteractiveLogonRight", "SeUndockPrivilege", "SeManageVolumePrivilege", "SeRemoteInteractiveLogonRight", "SeDenyRemoteInteractiveLogonRight", "SeImpersonatePrivilege", "SeCreateGlobalPrivilege")
	TASK08_REQUEST = Array("S-1-5-7","S-1-5-11","S-1-5-1","S-1-5-6","S-1-5-10","S-1-5-4","S-1-1-0","S-1-5-13")
	Dim TASK08_RETURN(8)
	Dim TEMP08
	FOR x = 0 TO ubound(TASK08)
		TEMP08 = ReadInI(SECPOL_FILENAME, "Privilege Rights", TASK08(x)) & ","
		
		FOR y = 0 TO ubound(TASK08_REQUEST)
			IF InStr(TEMP08, TASK08_REQUEST(y) & ",") <> 0 THEN TASK08_RETURN(y) = TASK08_RETURN(y) & TASK08(x) & NEWLINE_CHARACTER
		NEXT
		'IF InStr(TEMPVALUE,"S-1-5-7") <> 0 THEN TASK12_RETURN(0) = TASK12_RETURN(0) & NEWLINE_CHARACTER & TEMPVALUE 'ANONYMOUS LOGON	S-1-5-7
		'IF InStr(TEMPVALUE,"S-1-5-11") <> 0 THEN TASK12_RETURN(1) = TASK12_RETURN(1) & NEWLINE_CHARACTER & TEMPVALUE 'Authenticated Users	S-1-5-11
		'IF InStr(TEMPVALUE,"S-1-5-1") <> 0 THEN TASK12_RETURN(2) = TASK12_RETURN(2) & NEWLINE_CHARACTER & TEMPVALUE 'DIALUP	S-1-5-1
		'IF InStr(TEMPVALUE,"S-1-5-6") <> 0 THEN TASK12_RETURN(3) = TASK12_RETURN(3) & NEWLINE_CHARACTER & TEMPVALUE 'SERVICE	S-1-5-6
		'IF InStr(TEMPVALUE,"S-1-5-10") <> 0 THEN TASK12_RETURN(4) = TASK12_RETURN(4) & NEWLINE_CHARACTER & TEMPVALUE 'SELF
		'IF InStr(TEMPVALUE,"S-1-5-4") <> 0 THEN TASK12_RETURN(5) = TASK12_RETURN(5) & NEWLINE_CHARACTER & TEMPVALUE 'INTERACTIVE	S-1-5-4
		'IF InStr(TEMPVALUE,"S-1-1-0") <> 0 THEN TASK12_RETURN(6) = TASK12_RETURN(6) & NEWLINE_CHARACTER & TEMPVALUE 'Everyone	S-1-1-0
		'IF InStr(TEMPVALUE,"S-1-5-13") <> 0 THEN TASK12_RETURN(7) = TASK12_RETURN(7) & NEWLINE_CHARACTER & TEMPVALUE 'TERMINAL SERVER USER	S-1-5-13
	NEXT

	'FOR x = 0 TO ubound(TASK08_REQUEST)
	'	TASK08_REQUEST(x) = Replace(ReplaceSID(TASK08_REQUEST(x)),",",NEWLINE_CHARACTER)
	'NEXT
		
		
	IF objFSO.FileExists(SECPOL_FILENAME) THEN objFSO.DeleteFile SECPOL_FILENAME





	FOR x = 0 TO ubound(TASK31)
		IF TASK31_QUERY(X)= TASK31_REQUIRED(X) THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
		'WRITE_OUTPUT "3.1." & (x+1), TASK31(x), TASK31_QUERY(x), strCompare
		
		requireValue = TASK31_REQUIRED(X)
		currentValue = TASK31_QUERY(X)
		WRITE_OUTPUT_5 "3.1." & (x+1), TASK31(x), requireValue, currentValue, strCompare
		
	NEXT

WRITE_OUTPUT_SECTION_END "SECTION 3.1"

'================================================================================
'  SECTION 3.2
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 3.2 - Account Lockout Policy", "Local Security Settings > Account Policies > Account Lockout Policy"
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"
  


	FOR x = 0 TO ubound(TASK32)
		IF TASK32_QUERY(X)= TASK32_REQUIRED(X) THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
		'WRITE_OUTPUT "3.2." & (x+1), TASK32(x), TASK32_QUERY(x), strCompare
		
		requireValue = TASK32_REQUIRED(X)
		currentValue = TASK32_QUERY(X)
		WRITE_OUTPUT_5 "3.2." & (x+1), TASK32(x), requireValue, currentValue, strCompare
	NEXT


WRITE_OUTPUT_SECTION_END "SECTION 3.2"


'================================================================================
'  SECTION 3.3
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 3.3 - User Rights Assignment", "Local Security Settings > Local Policies > User Right Assignment"
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"
  


Function CheckTASK33(ByVal x, ByVal yield)
	IF NOT (x=14 OR x=15) THEN
		ERROR_MSG = "X"
		IF InStr(yield,"SUPPORT_") <> 0 THEN ERROR_MSG = ERROR_MSG & " (SUPPORT FOUND)"
		IF InStr(yield,"CUSTOMGUEST") <> 0 THEN ERROR_MSG = ERROR_MSG & " (CUSTOMGUEST FOUND)"
		IF InStr(yield,"*S-1-5-7") <> 0 THEN ERROR_MSG = ERROR_MSG & " (ANONYMOUS FOUND)"
		IF InStr(yield,"*S-1-5-32-546") <> 0 THEN ERROR_MSG = ERROR_MSG & " (GUESTS GRP FOUND)"
		IF NOT ERROR_MSG = "X" THEN CheckTASK33 = ERROR_MSG : EXIT FUNCTION
	END IF
	
	SELECT CASE (x+1)
	CASE 1 : IF InStr(yield,"*S-1-5-11") = 0 THEN CheckTASK33 = "NO AUTHENTICATE USER " ELSE CheckTASK33 = CONFIRM_CHARACTER() 
	CASE 2 : IF yield = "" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "NOT BLANK"
	CASE 3 :
		IF yield = "" THEN
			CheckTASK33 = CONFIRM_CHARACTER() & " NO AUTHENTICATE USER"
		ELSEIF InStr(yield,"*S-1-5-11") = 0 THEN
			CheckTASK33 = ERROR_CHARACTER() & " NO AUTHENTICATE USER"
		ELSE
			CheckTASK33 = CONFIRM_CHARACTER() 
		END IF
	CASE 4 : CheckTASK33 = CONFIRM_CHARACTER()
	CASE 5 : CheckTASK33 = CONFIRM_CHARACTER()
	CASE 6 : CheckTASK33 = CONFIRM_CHARACTER()
	CASE 7 : CheckTASK33 = CONFIRM_CHARACTER()
	CASE 8 : IF InStr(yield,"*S-1-5-11") = 0 THEN CheckTASK33 = "NO AUTHENTICATE USER, " ELSE CheckTASK33 = CONFIRM_CHARACTER() 
	CASE 9 : IF yield = "*S-1-5-32-544,*S-1-5-19" OR yield = "*S-1-5-19,*S-1-5-32-544" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "MORE THAN TWO"
	CASE 10 : IF yield = "*S-1-5-32-544" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "REQUIRE SOLE-ADMIN, "
	CASE 11 : IF yield = "" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "NOT BLANK, "
	CASE 12 : IF yield = "*S-1-5-32-544,*S-1-5-6" OR yield = "*S-1-5-6,*S-1-5-32-544" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "MORE THAN TWO"
	CASE 13 : IF yield = "" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "NOT BLANK, "
	CASE 14 : IF yield = "*S-1-5-32-544" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "REQUIRE SOLE-ADMIN, "
	CASE 15
		ERROR_MSG = ""
		IF InStr(yield,"SUPPORT_") = 0 THEN ERROR_MSG = ERROR_MSG & " NO SUPPORT, "
		IF InStr(yield,"CUSTOMGUEST") = 0 THEN ERROR_MSG = ERROR_MSG & " NO CUSTOMGUEST, "
		IF InStr(yield,"*S-1-5-7") = 0 THEN ERROR_MSG = ERROR_MSG & " NO ANONYMOUS, "
		IF InStr(yield,"*S-1-5-32-546") = 0 THEN ERROR_MSG = ERROR_MSG & " NO GUESTS GRP, "
		IF ERROR_MSG = "" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = ERROR_MSG
		'IF ERROR_MSG = "" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = ERROR_CHARACTER() & MID(ERROR_MSG,2)
	CASE 16
		ERROR_MSG = ""
		IF InStr(yield,"SUPPORT_") = 0 THEN ERROR_MSG = ERROR_MSG & "NO SUPPORT, "
		IF InStr(yield,"CUSTOMGUEST") = 0 THEN ERROR_MSG = ERROR_MSG & "NO CUSTOMGUEST, "
		IF InStr(yield,"*S-1-5-32-546") = 0 THEN ERROR_MSG = ERROR_MSG & "NO GUESTS GRP, "
		IF ERROR_MSG = "" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = ERROR_MSG
		'IF ERROR_MSG = "" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = ERROR_CHARACTER() & MID(ERROR_MSG,2)
	CASE 17 : IF yield = "*S-1-5-32-544" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "REQUIRE SOLE-ADMIN, "
	CASE 18 : IF yield = "*S-1-5-32-544" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "REQUIRE SOLE-ADMIN, "
	CASE 19 : IF yield = "*S-1-5-32-544" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "REQUIRE SOLE-ADMIN, "
	CASE 20 : IF yield = "*S-1-5-32-544" THEN CheckTASK33 = CONFIRM_CHARACTER() ELSE CheckTASK33 = "REQUIRE SOLE-ADMIN, "
	
	CASE ELSE
	CheckTASK33 = "OUT OF CHECKING BOUND"
	END SELECT
	
	IF (NOT CheckTASK33 = CONFIRM_CHARACTER()) AND (NOT x=2) THEN CheckTASK33 = ERROR_CHARACTER() & " " & CheckTASK33
	
End Function

TASK33_REQUIRED = Array( _
"*S-1-5-11,*S-1-5-32-544", _
"", _
"", _
"*S-1-5-19,*S-1-5-20,*S-1-5-32-544", _
"*S-1-5-32-544", _
"*S-1-5-32-544", _
"*S-1-5-32-544", _
"*S-1-5-11,*S-1-5-32-544", _
"*S-1-5-19,*S-1-5-32-544", _
"*S-1-5-32-544", _
"", _
"*S-1-5-32-544,*S-1-5-6", _
"", _
"*S-1-5-32-544", _
"SUPPORT_AABBCCDDEE,CUSTOMGUEST,*S-1-5-32-546,*S-1-5-7", _
"SUPPORT_AABBCCDDEE,CUSTOMGUEST,*S-1-5-32-546", _
"*S-1-5-32-544", _
"*S-1-5-32-544", _
"*S-1-5-32-544", _
"*S-1-5-32-544" _
)
	FOR x = 0 TO ubound(TASK33)
		strCompare = CheckTASK33(x, TASK33_QUERY(x))
		
		requireValue = "<div>" & Replace(ReplaceSID(TASK33_REQUIRED(x)),",",NEWLINE_CHARACTER) & "</div>"
		currentValue = "<div>" & Replace(ReplaceSID(TASK33_QUERY(x)),",",NEWLINE_CHARACTER) & "</div>"
		WRITE_OUTPUT_5 "3.3." & (x+1), TASK33_DESC(x), requireValue, currentValue, strCompare
	NEXT
WRITE_OUTPUT_SECTION_END "SECTION 3"


'================================================================================
'  SECTION 4
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 4 - Security Options", "Local Security Settings > Local Policies > Security Options"
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"

ITEM_LABEL = Array( _ 
"Accounts: Administrator account status", _ 
"Accounts: Guest account status", _ 
"Accounts: Limit local account use of blank passwords to console logon only", _ 
"Accounts: Rename administrator account", _ 
"Accounts: Rename guest account", _ 
"Audit: Audit the access of global system objects", _ 
"Audit: Audit the use of Backup and Restore privilege", _ 
"Audit: Shut down system immediately if unable to log security audits", _ 
"DCOM: Machine Access Restrictions in Security Descriptor Definition Language (SDDL) syntax", _ 
"DCOM: Machine Launch Restrictions in Security Descriptor Definition Language (SDDL) syntax", _ 
"Devices: Allow undock without having to log on", _ 
"Devices: Allowed to format and eject removable media", _ 
"Devices: Prevent users from installing printer drivers", _ 
"Devices: Restrict CD-ROM access to locally logged-on user only", _ 
"Devices: Restrict floppy access to locally logged-on user only", _ 
"Devices: Unsigned driver installation behavior", _ 
"Domain controller: Allow server operators to schedule tasks", _ 
"Domain controller: LDAP server signing requirements", _ 
"Domain controller: Refuse machine account password changes", _ 
"Domain member: Digitally encrypt or sign secure channel data (always)", _ 
"Domain member: Digitally encrypt secure channel data (when possible)", _ 
"Domain member: Digitally sign secure channel data (when possible)", _ 
"Domain member: Disable machine account password changes", _ 
"Domain member: Maximum machine account password age", _ 
"Domain member: Require strong (Windows 2000 or later) session key", _ 
"Interactive logon: Display user information when the session is locked", _ 
"Interactive logon: Do not display last user name", _ 
"Interactive logon: Do not require CTRL+ALT+DEL", _ 
"Interactive logon: Message text for users attempting to log on", _ 
"Interactive logon: Message title for users attempting to log on", _ 
"Interactive logon: Number of previous logons to cache (in case domain controller is not available)", _ 
"Interactive logon: Prompt user to change password before expiration", _ 
"Interactive logon: Require Domain Controller authentication to unlock workstation", _ 
"Interactive logon: Require smart card", _ 
"Interactive logon: Smart card removal behavior", _ 
"Microsoft network client: Digitally sign communications (always)", _ 
"Microsoft network client: Digitally sign communications (if server agrees)", _ 
"Microsoft network client: Send unencrypted password to third-party SMB servers", _ 
"Microsoft network server: Amount of idle time required before suspending session", _ 
"Microsoft network server: Digitally sign communications (always)", _ 
"Microsoft network server: Digitally sign communications (if client agrees)", _ 
"Microsoft network server: Disconnect clients when logon hours expire", _ 
"Network access: Allow anonymous SID/Name translation", _ 
"Network access: Do not allow anonymous enumeration of SAM accounts", _ 
"Network access: Do not allow anonymous enumeration of SAM accounts and shares", _ 
"Network access: Do not allow storage of credentials or .NET Passports for network authentication", _ 
"Network access: Let Everyone permissions apply to anonymous users", _ 
"Network access: Named Pipes that can be accessed anonymously", _ 
"Network access: Remotely accessible registry paths", _ 
"Network access: Remotely accessible registry paths and sub-paths", _ 
"Network access: Restrict anonymous access to Named Pipes and Shares", _ 
"Network access: Shares that can be accessed anonymously", _ 
"Network access: Sharing and security model for local accounts", _ 
"Network security: Do not store LAN Manager hash value on next password change", _ 
"Network security: Force logoff when logon hours expire", _ 
"Network security: LAN Manager authentication level", _ 
"Network security: LDAP client signing requirements", _ 
"Network security: Minimum session security for NTLM SSP based (including secure RPC) clients", _ 
"Network security: Minimum session security for NTLM SSP based (including secure RPC) servers", _ 
"Recovery console: Allow automatic administrative logon", _ 
"Recovery console: Allow floppy copy and access to all drives and all folders", _ 
"Shutdown: Allow system to be shut down without having to log on", _ 
"Shutdown: Clear virtual memory pagefile", _ 
"System cryptography: Force strong key protection for user keys stored on the computer", _ 
"System cryptography: Use FIPS compliant algorithms for encryption, hashing, and signing", _ 
"System objects: Default owner for objects created by members of the Administrators group", _ 
"System objects: Require case insensitivity for non-Windows subsystems", _ 
"System objects: Strengthen default permissions of internal system objects (e.g. Symbolic Links)", _ 
"System settings: Optional subsystems", _ 
"System settings: Use Certificate Rules on Windows Executables for Software Restriction Policies" _
)

ITEM_REG = Array( _
"", _
"", _
"System/CurrentControlSet/Control/Lsa/LimitBlankPasswordUse", _
"", _
"", _
"System/CurrentControlSet/Control/Lsa/AuditBaseObjects", _
"System/CurrentControlSet/Control/Lsa/FullPrivilegeAuditing", _
"System/CurrentControlSet/Control/Lsa/CrashOnAuditFail", _
"SOFTWARE/policies/Microsoft/windows NT/DCOM/MachineAccessRestriction", _
"SOFTWARE/policies/Microsoft/windows NT/DCOM/MachineLaunchRestriction", _
"Software/Microsoft/Windows/CurrentVersion/Policies/System/UndockWithoutLogon", _
"Software/Microsoft/Windows NT/CurrentVersion/Winlogon/AllocateDASD", _
"System/CurrentControlSet/Control/Print/Providers/LanMan Print Services/Servers/AddPrinterDrivers", _
"Software/Microsoft/Windows NT/CurrentVersion/Winlogon/AllocateCDRoms", _
"Software/Microsoft/Windows NT/CurrentVersion/Winlogon/AllocateFloppies", _
"Software/Microsoft/Driver Signing/Policy", _
"System/CurrentControlSet/Control/Lsa/SubmitControl", _
"System/CurrentControlSet/Services/NTDS/Parameters/LDAPServerIntegrity", _
"System/CurrentControlSet/Services/Netlogon/Parameters/RefusePasswordChange", _
"System/CurrentControlSet/Services/Netlogon/Parameters/RequireSignOrSeal", _
"System/CurrentControlSet/Services/Netlogon/Parameters/SealSecureChannel", _
"System/CurrentControlSet/Services/Netlogon/Parameters/SignSecureChannel", _
"System/CurrentControlSet/Services/Netlogon/Parameters/DisablePasswordChange", _
"System/CurrentControlSet/Services/Netlogon/Parameters/MaximumPasswordAge", _
"System/CurrentControlSet/Services/Netlogon/Parameters/RequireStrongKey", _
"Software/Microsoft/Windows/CurrentVersion/Policies/System/DontDisplayLockedUserId", _
"Software/Microsoft/Windows/CurrentVersion/Policies/System/DontDisplayLastUserName", _
"Software/Microsoft/Windows/CurrentVersion/Policies/System/DisableCAD", _
"Software/Microsoft/Windows/CurrentVersion/Policies/System/LegalNoticeText", _
"Software/Microsoft/Windows/CurrentVersion/Policies/System/LegalNoticeCaption", _
"Software/Microsoft/Windows NT/CurrentVersion/Winlogon/CachedLogonsCount", _
"Software/Microsoft/Windows NT/CurrentVersion/Winlogon/PasswordExpiryWarning", _
"Software/Microsoft/Windows NT/CurrentVersion/Winlogon/ForceUnlockLogon", _
"Software/Microsoft/Windows/CurrentVersion/Policies/System/ScForceOption", _
"Software/Microsoft/Windows NT/CurrentVersion/Winlogon/ScRemoveOption", _
"System/CurrentControlSet/Services/LanmanWorkstation/Parameters/RequireSecuritySignature", _
"System/CurrentControlSet/Services/LanmanWorkstation/Parameters/EnableSecuritySignature", _
"System/CurrentControlSet/Services/LanmanWorkstation/Parameters/EnablePlainTextPassword", _
"System/CurrentControlSet/Services/LanManServer/Parameters/AutoDisconnect", _
"System/CurrentControlSet/Services/LanManServer/Parameters/RequireSecuritySignature", _
"System/CurrentControlSet/Services/LanManServer/Parameters/EnableSecuritySignature", _
"System/CurrentControlSet/Services/LanManServer/Parameters/EnableForcedLogOff", _
"", _
"System/CurrentControlSet/Control/Lsa/RestrictAnonymousSAM", _
"System/CurrentControlSet/Control/Lsa/RestrictAnonymous", _
"System/CurrentControlSet/Control/Lsa/DisableDomainCreds", _
"System/CurrentControlSet/Control/Lsa/EveryoneIncludesAnonymous", _
"System/CurrentControlSet/Services/LanManServer/Parameters/NullSessionPipes", _
"System/CurrentControlSet/Control/SecurePipeServers/Winreg/AllowedExactPaths/Machine", _
"System/CurrentControlSet/Control/SecurePipeServers/Winreg/AllowedPaths/Machine", _
"System/CurrentControlSet/Services/LanManServer/Parameters/RestrictNullSessAccess", _
"System/CurrentControlSet/Services/LanManServer/Parameters/NullSessionShares", _
"System/CurrentControlSet/Control/Lsa/ForceGuest", _
"System/CurrentControlSet/Control/Lsa/NoLMHash", _
"", _
"System/CurrentControlSet/Control/Lsa/LmCompatibilityLevel", _
"System/CurrentControlSet/Services/LDAP/LDAPClientIntegrity", _
"System/CurrentControlSet/Control/Lsa/MSV1_0/NTLMMinClientSec", _
"System/CurrentControlSet/Control/Lsa/MSV1_0/NTLMMinServerSec", _
"Software/Microsoft/Windows NT/CurrentVersion/Setup/RecoveryConsole/SecurityLevel", _
"Software/Microsoft/Windows NT/CurrentVersion/Setup/RecoveryConsole/SetCommand", _
"Software/Microsoft/Windows/CurrentVersion/Policies/System/ShutdownWithoutLogon", _
"System/CurrentControlSet/Control/Session Manager/Memory Management/ClearPageFileAtShutdown", _
"Software/Policies/Microsoft/Cryptography/ForceKeyProtection", _
"System/CurrentControlSet/Control/Lsa/FIPSAlgorithmPolicy", _
"System/CurrentControlSet/Control/Lsa/NoDefaultAdminOwner", _
"System/CurrentControlSet/Control/Session Manager/Kernel/ObCaseInsensitive", _
"System/CurrentControlSet/Control/Session Manager/ProtectionMode", _
"System/CurrentControlSet/Control/Session Manager/SubSystems/optional", _
"Software/Policies/Microsoft/Windows/Safer/CodeIdentifiers/AuthenticodeEnabled" _
)

ITEM_REQ = Array( _
"1", _
"0", _
"1", _
"CUSTOMADMIN", _
"CUSTOMGUEST", _
"0", _
"0", _
"0", _
"Not Defined", _
"Not Defined", _
"0", _
"0", _ 
"1", _
"1", _
"1", _
"1", _ 
"0", _
"Not Defined", _
"0", _
"1", _
"1", _
"1", _
"0", _
"30", _
"1", _
"2", _
"1", _
"0", _
"This system is restricted to authorised users. If unauthorised, terminate access now! Clicking OK indicates your acceptance of the information in the background.", _
"IT IS AN OFFENSE TO CONTINUE WITHOUT PROPER AUTHORISATION", _
"0", _
"14", _
"1", _
"0", _
"0", _
"0", _
"1", _
"0", _
"15", _ 
"1", _
"1", _
"1", _
"0", _
"1", _
"1", _
"1", _
"0", _
"", _ 
"System\CurrentControlSet\Control\ProductOptions | System\CurrentControlSet\Control\Server Applications | Software\Microsoft\Windows NT\CurrentVersion | ", _
"System\CurrentControlSet\Control\Print\Printers | System\CurrentControlSet\Services\Eventlog | Software\Microsoft\OLAP Server | Software\Microsoft\Windows NT\CurrentVersion\Print | Software\Microsoft\Windows NT\CurrentVersion\Windows | System\CurrentControlSet\Control\ContentIndex | System\CurrentControlSet\Control\Terminal Server | System\CurrentControlSet\Control\Terminal Server\UserConfig | System\CurrentControlSet\Control\Terminal Server\DefaultUserConfiguration | Software\Microsoft\Windows NT\CurrentVersion\Perflib | System\CurrentControlSet\Services\SysmonLog | ", _
"1", _
"", _
"0", _
"1", _
"0", _
"5", _ 
"1", _ 
"537395248", _
"537395248", _ 
"0", _
"0", _
"0", _
"1", _
"Not Defined", _
"0", _
"0", _ 
"1", _
"1", _
"", _
"1" _
)

ITEM_BINARY = Array( _
"1", _
"1", _
"1", _
"0", _ 
"0", _ 
"1", _
"1", _
"1", _
"0", _ 
"0", _ 
"1", _
"0", _ 
"1", _
"1", _
"1", _
"0", _ 
"1", _
"0", _ 
"1", _
"1", _
"1", _
"1", _
"1", _
"0", _ 
"1", _
"0", _
"1", _
"1", _
"0", _ 
"0", _ 
"0", _
"0", _
"1", _
"1", _
"0", _
"1", _
"1", _
"1", _
"0", _ 
"1", _
"1", _
"1", _
"1", _
"1", _
"1", _
"1", _
"1", _
"0", _ 
"0", _ 
"0", _ 
"1", _
"0", _ 
"0", _
"1", _
"1", _
"0", _ 
"0", _ 
"0", _ 
"0", _ 
"1", _
"1", _
"1", _
"1", _
"0", _ 
"1", _
"0", _ 
"1", _
"1", _
"0", _
"1" _
)

BINARY_VALUE = ARRAY("Disabled","Enabled")
ITEM_REQ_SEC = ARRAY("1", "0", "", "CUSTOMADMIN", "CUSTOMGUEST", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "0", "", "", "", "", "", "", "", "", "", "", "", "0", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

SECEDIT_PATH = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SeCEdit\Reg Values\MACHINE/"

FOR x = 0 TO ubound(ITEM_LABEL)
	
	strAuditNo = "4." & (X+1)
	'wscript.stdout.write FPadding("4." & (X+1), 8) & FPadding(ITEM_LABEL(X), 49) & "  "
	IF ITEM_REG(X) <> "" THEN
	
		strKeyPath = SECEDIT_PATH & ITEM_REG(X)
		keysize = Instr(StrReverse(strKeyPath),"/")
		ACTUAL_KEY_PATH = Mid(Replace(ITEM_REG(X),"/","\"), 1, len(ITEM_REG(X)) - keysize)
		ACTUAL_KEY     = Mid(strKeyPath, len(strKeyPath) - keysize + 2)
		
		IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, "ValueType", value) <> 0 Then
			value = "Not Found"
		ELSE
			
			SELECT CASE value
			
			case 1 'REG_SZ
				IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, ACTUAL_KEY_PATH, ACTUAL_KEY, value) <> 0 Then value="Not Found"
			case 2 'REG_EXPAND_SZ  // with environment variables to expand
				value = "XXX"
			case 3 'REG_BINARY
				IF oReg.GetBinaryValue(HKEY_LOCAL_MACHINE, ACTUAL_KEY_PATH, ACTUAL_KEY, Bvalues) <> 0 Then
					value = "Not Found"
				ELSE
					value=""
					For i = lBound(Bvalues) to uBound(Bvalues)
						value = value & Bvalues(i)
					Next

				END IF
			case 4 'REG_DWORD
				IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, ACTUAL_KEY_PATH, ACTUAL_KEY, value) <> 0 Then value="Not Found"
			case 7 'REG_MULTI_SZ
				IF oReg.GetMultiStringValue(HKEY_LOCAL_MACHINE, ACTUAL_KEY_PATH, ACTUAL_KEY, arrValues) <> 0 Then
					IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, ACTUAL_KEY_PATH, ACTUAL_KEY, value) <> 0 Then
					value = "Not Found"
					END IF
				ELSE
					value = ""
					For Each strValue in arrValues
					IF NOT strValue="" THEN value = value & strValue & " | "
					Next
				END IF
	
			case else
				wscript.quit "IMPOSSIBLE ERROR"
			END SELECT
		END IF
		
		IF StrComp(ITEM_REQ(X),value)=0 THEN
			strCompare = CONFIRM_CHARACTER()
		ELSEIF ITEM_REQ(X) = "Not Defined" AND value = "Not Found" THEN	
			strCompare = CONFIRM_CHARACTER()
		ELSEIF value = "Not Found" THEN
			strCompare = NA_CHARACTER() & " NOT FOUND"
		ELSEIF value = "N/A" THEN
			strCompare = NA_CHARACTER()
		ELSE
			strCompare = ERROR_CHARACTER()
		END IF
		
		IF len(strCompare) >= 38 AND HTML_MODE = FALSE THEN strCompare = FPadding(strCompare, 35) & "..."
		
		'wscript.stdout.write FPadding(value, 8) & ITEM_LABEL(X) '& ITEM_REG(X)
		'wscript.stdout.write FPadding(value, 20) & FPadding(ITEM_REQ(X), 20) & " " & strCompare
		'wscript.stdout.write FPadding(value, 20) & " " & strCompare
	ELSE
		
		SELECT CASE (X+1)
			CASE 1: value = task_41
			CASE 2: value = task_42
			CASE 4: value = task_44
			CASE 5: value = task_45
			CASE 43: value = task_443
			CASE 55: value = task_455
			CASE ELSE
			wscript.stdout.write FPadding(ITEM_REQ_SEC(x), 20) & " " & "MANUAL 4." & (X+1)
		END SELECT
		
		IF StrComp(ITEM_REQ(X),value)=0 THEN
			strCompare = CONFIRM_CHARACTER()
		ELSEIF value = "[NOT FOUND]" THEN
			strCompare = NA_CHARACTER() & " NOT FOUND"
		ELSEIF value = "N/A" THEN
			strCompare = NA_CHARACTER()
		ELSE
			strCompare = ERROR_CHARACTER()
		END IF


		'wscript.stdout.write FPadding(value, 20) & " " & strCompare
		
		
		'wscript.stdout.write FPadding("", 40) & " " & "4." & (X+1) & " MANUAL"
		'wscript.stdout.write FPadding(ITEM_REQ_SEC(x), 20) & " " & "MANUAL 4." & (X+1)
		
		
		'wscript.stdout.write FPadding("=====" & "4." & (X+1) & "=MANUAL=("&ITEM_REQ_SEC(x)&")=====", 40)
	END IF
	
	IF ITEM_BINARY(X)="1" THEN
		ITEM_REQ(X) = BINARY_VALUE(CInt(ITEM_REQ(X)))
		IF NOT value = "Not Found" THEN value = BINARY_VALUE(CInt(value))
	ELSE
	
		IF oReg.GetMultiStringValue(HKEY_LOCAL_MACHINE, strKeyPath, "DisplayChoices", arrValues) = 0 Then
			For Each strValue in arrValues
				subValue = Split(strValue, "|")
				IF CStr(subValue(0)) = CStr(Left(value,1)) THEN value = subValue(1)
				IF CStr(subValue(0)) = CStr(Left(ITEM_REQ(X),1)) THEN ITEM_REQ(X) = subValue(1)
			Next
		END IF

	END IF
	
	IF HTML_MODE = TRUE THEN strCurrent = "<div>" & value & "</div>"
	IF HTML_MODE = TRUE THEN strRequire = "<div>" & ITEM_REQ(X) & "</div>"
	IF (X+1)=49 OR (X+1)=50 THEN
		'wscript.stderr.write "X:" & X & vbcrlf
		'wscript.stderr.write "ITEM_REQ(X):" & ITEM_REQ(X) & vbcrlf
		'wscript.stderr.write "value:" & value & vbcrlf
		'wscript.stderr.write vbcrlf
	
		ORIGINAL_X = X
		arrNOTFULFILLED = arrayFilter(split(ITEM_REQ(X), " | "), split(value, " | "), arrADDITIONAL)
		X = ORIGINAL_X
		
		IF Len(Join(arrNOTFULFILLED, NEWLINE_CHARACTER()))=0 AND Len(Join(arrADDITIONAL, NEWLINE_CHARACTER()))=0 THEN
			strCompare = CONFIRM_CHARACTER()
		ELSEIF Len(Join(arrNOTFULFILLED, NEWLINE_CHARACTER()))=0 THEN
			IF (X+1)=50 AND (NOT(Cint(VALUE_HTTPSSL) = 4)) AND StrComp(UCase(CStr(Join(arrADDITIONAL, ""))),UCase("SYSTEM\CurrentControlSet\Services\CertSvc"),1)=0 THEN
				strCompare = CONFIRM_CHARACTER() & "CertSvc - HTTPSSL"
			ELSE
				strCompare = WARNING_CHARACTER() & vbcrlf
				strCompare = strCompare & "<b>ADDITIONAL:</b>" & vbcrlf & Join(arrADDITIONAL, NEWLINE_CHARACTER())
			END IF
		ELSE
				strCompare = ERROR_CHARACTER() & vbcrlf
				IF NOT Len(Join(arrNOTFULFILLED, NEWLINE_CHARACTER()))=0 THEN strCompare = strCompare & "<b>NOT FULFILLED:</b>" & vbcrlf & Join(arrNOTFULFILLED, NEWLINE_CHARACTER())
				IF NOT Len(Join(arrADDITIONAL, NEWLINE_CHARACTER()))=0 THEN strCompare = strCompare & "<b>ADDITIONAL:</b>" & vbcrlf & Join(arrADDITIONAL, NEWLINE_CHARACTER())
		END IF
		
		WRITE_OUTPUT_5 strAuditNo, ITEM_LABEL(X), Replace(strRequire," | ",NEWLINE_CHARACTER), Replace(strCurrent," | ",NEWLINE_CHARACTER), strCompare
	ELSE
		WRITE_OUTPUT_5 strAuditNo, ITEM_LABEL(X), strRequire, strCurrent, strCompare
	END IF



	'wscript.stdout.write vbcrlf
NEXT
WRITE_OUTPUT_SECTION_END "SECTION 4"


'================================================================================
'  SECTION 4 (RECONFIRM)
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 4 RECONFIRM - Security Options", "Registry Checking. Besides, some values can only be tuned from registry."
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"

	strAuditNo = "4.1"
	strKeyPath = "SYSTEM\CurrentControlSet\Control\LSA"
	strKeyName = "RestrictAnonymous"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	strCurrent=value
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, strCurrent, strCompare

	strAuditNo = "4.2"
	strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
	strKeyName = "ShutDownWithoutLogon"
	strRequire = "0"
	IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	strCurrent=value
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, strCurrent, strCompare

	strAuditNo = "4.3"
	strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"
	strKeyName = "ClearPageFileAtShutdown"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	strCurrent=value
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, strCurrent, strCompare

	strAuditNo = "4.4"
	strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
	strKeyName = "DontDisplayLastUserName"
	strRequire = "1"
	IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	strCurrent=value
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, strCurrent, strCompare

	strAuditNo = "4.5"
	strKeyPath = "SYSTEM\CurrentControlSet\Control\LSA"
	strKeyName = "LMCompatibilityLevel"
	strRequire = 5
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	strCurrent=value
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, strCurrent, strCompare

	strAuditNo = "4.6"
	strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
	strKeyName = "CachedLogonsCount"
	strRequire = "0"
	IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	strCurrent=value
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, strCurrent, strCompare

	strAuditNo = "4.7"
	strKeyPath = "SYSTEM\CurrentControlSet\Control\Print\Providers\LanMan Print Services\Servers\"
	strKeyName = "AddPrinterDrivers"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	strCurrent=value
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, strCurrent, strCompare

	strAuditNo = "4.8"
	strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
	strKeyName = "AllocateCDRoms"
	strRequire = "1"
	IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	strCurrent=value
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, strCurrent, strCompare
		 
	strAuditNo = "4.9"
	strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
	strKeyName = "AllocateFloppies"
	strRequire = "1"
	IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	strCurrent=value
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, strCurrent, strCompare

	strAuditNo = "4.10"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\LanManWorkstation\Parameters"
	strKeyName = "EnablePlainTextPassword"
	strRequire = 0
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	strCurrent=value
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, strCurrent, strCompare
WRITE_OUTPUT_SECTION_END "SECTION 4 (RECONFIRM)"




'================================================================================
'  SECTION 5
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 5 - Event Logging", ""
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"

	FOR x = 0 TO ubound(TASK5)
		IF TASK5_QUERY(X)= TASK5_REQUIRED(X) THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
		WRITE_OUTPUT_5 "5." & (x+1), TASK5(x), EVENT_STATUS(TASK5_REQUIRED(X)), EVENT_STATUS(TASK5_QUERY(x)), strCompare
	NEXT

	strAuditNo = "5.10"
	strDesc = "MaxSize (System)"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\Eventlog\System"
	strKeyName = "MaxSize"
	strRequire = "1024000KB"
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED" ELSE value = (ReinterpretSignedAsUnsigned(value)/1024) & "KB"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strDesc, strRequire, value, strCompare

	strAuditNo = "5.10"
	strDesc = "MaxSize (Application)"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\Eventlog\Application"
	strKeyName = "MaxSize"
	strRequire = "1024000KB"
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED" ELSE value = (ReinterpretSignedAsUnsigned(value)/1024) & "KB"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strDesc, strRequire, value, strCompare

	strAuditNo = "5.11"
	strDesc = "MaxSize (Security)"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\Eventlog\Security"
	strKeyName = "MaxSize"
	'strRequire = "4194240KB"
	strRequire = "1024000KB"
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED" ELSE value = (ReinterpretSignedAsUnsigned(value)/1024) & "KB"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strDesc, strRequire, value, strCompare

	strAuditNo = "5.12"
	strDesc = "Retention (System)"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\Eventlog\System"
	strKeyName = "Retention"
	strRequire = "30DAYS"
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED" ELSE value = CInt(value/86400) & "DAYS"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strDesc, strRequire, value, strCompare

	strAuditNo = "5.12"
	strDesc = "Retention (Application)"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\Eventlog\Application"
	strKeyName = "Retention"
	strRequire = "30DAYS"
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED" ELSE value = CInt(value/86400) & "DAYS"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strDesc, strRequire, value, strCompare

	strAuditNo = "5.12"
	strDesc = "Retention (Security)"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\Eventlog\Security"
	strKeyName = "Retention"
	strRequire = "30DAYS"
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED" ELSE value = CInt(value/86400) & "DAYS"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strDesc, strRequire, value, strCompare

	strAuditNo = "5.13"
	strDesc = "RestrictGuestAccess (Application)"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\EventLog\Application"
	strKeyName = "RestrictGuestAccess"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strDesc, strRequire, value, strCompare

	strAuditNo = "5.13"
	strDesc = "RestrictGuestAccess (Security)"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\EventLog\Security"
	strKeyName = "RestrictGuestAccess"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strDesc, strRequire, value, strCompare

	strAuditNo = "5.13"
	strDesc = "RestrictGuestAccess (System)"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\EventLog\System"
	strKeyName = "RestrictGuestAccess"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strDesc, strRequire, value, strCompare

	strAuditNo = "5.14"
	strDesc = "WarningLevel (Security)"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\EventLog\Security"
	strKeyName = "WarningLevel"
	strRequire = 90
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strDesc, strRequire, value, strCompare

WRITE_OUTPUT_SECTION_END "SECTION 5"


'================================================================================
'  SECTION 6
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 6 - Miscellaneous Registry Settings", ""
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"

	strAuditNo = "6.1"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\RasMan\Parameters"
	strKeyName = "DisableSavePassword"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.1.1"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\RasMan\PPP"
	strKeyName = "ForceEncryptedData"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.1.2"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\RasMan\Parameters"
	strKeyName = "Logging"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.1.3"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\RasMan\PPP"
	strKeyName = "SecureVPN"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.1.4"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\RasMan\PPP"
	strKeyName = "ForceEncryptedPassword"
	strRequire = 2
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.2"
	strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager\SubSystems"
	strKeyName = "Optional"
	strRequire = " (REG_MULTI_SZ)"
	'IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF oReg.GetMultiStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, arrValues) <> 0 Then
		IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
		value = value & " (REG_SZ)"
	ELSE
		value = ""
		For Each strValue in arrValues
			IF NOT strValue="" THEN value = value & strValue & " | "
		Next
		value = value & " (REG_MULTI_SZ)"
	END IF
	
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

				
				
				
	strAuditNo = "6.3"
	strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager\SubSystems"
	strKeyName = "POSIX"
	strRequire = "NOT EXISTED"
	IF oReg.GetExpandedStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF value=strRequire THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.4"
	strKeyPath = "Software\Microsoft\OLE"
	strKeyName = "EnableDcom"
	strRequire = "N"
	IF oReg.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.5"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\LanManServer\Parameters"
	strKeyName = "NullSessionPipes"
	strRequire = ""
	IF oReg.GetMultiStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, arrValues) <> 0 Then
		value = "NOT EXISTED"
	ELSE
		value = ""
		For Each strValue in arrValues
			IF NOT value="" THEN value = value & strValue & " | "
		Next
	END IF
	IF value=strRequire THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.5"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\LanManServer\Parameters"
	strKeyName = "NullSessionShares"
	strRequire = ""
	IF oReg.GetMultiStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, arrValues) <> 0 Then
		value = "NOT EXISTED"
	ELSE
		value = ""
		For Each strValue in arrValues
			IF NOT value="" THEN value = value & strValue & " | "
		Next
	END IF
	IF value=strRequire THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.6"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\LanManServer\Parameters"
	strKeyName = "RestrictNullSessAccess"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.7"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\LanManServer\Parameters"
	strKeyName = "AutoShareServer"
	strRequire = 0
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.7"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\LanManServer\Parameters"
	strKeyName = "AutoShareWks"
	strRequire = 0
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.8"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\CDrom"
	strKeyName = "Autorun"
	strRequire = 0
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.9"
	strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
	strKeyName = "NoDriveTypeAutorun"
	strRequire = 255
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "6.9.1"
	strKeyPath = ".DEFAULT\Control Panel\Desktop"
	strKeyName = "ScreenSaverIsSecure"
	strRequire = "1"
	IF oReg.GetStringValue(HKEY_USERS, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare
WRITE_OUTPUT_SECTION_END "SECTION 6"

'================================================================================
'  SECTION 7
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 7 - Local Group Accounts", ""
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"


REQ_71 = "DO NOT add non-administrative accounts to this group."
REQ_72 = "DO NOT add non-administrative accounts to this group."
REQ_73 = "DO NOT add non-administrative accounts to this group." & NEWLINE_CHARACTER & "DO NOT assign any permission to this group."
REQ_74 = "DO NOT use this group." & NEWLINE_CHARACTER & "Remove all accounts." & NEWLINE_CHARACTER & "DO NOT assign permission"
REQ_75 = "DO NOT add non-administrative accounts to this group."
REQ_76 = "DO NOT add guest to this group."

Set colGroups = GetObject("WinNT://.")
colGroups.Filter = Array("group")
For Each objGroup In colGroups
	'Wscript.Echo objGroup.Name
	IF objGroup.Name = "Administrators" THEN For Each objUser in objGroup.Members : RET_71 = RET_71 & objUser.Name & NEWLINE_CHARACTER: Next
	IF objGroup.Name = "Backup Operators" THEN For Each objUser in objGroup.Members : RET_72 = RET_72 & objUser.Name & NEWLINE_CHARACTER: Next
	IF objGroup.Name = "Power Users" THEN For Each objUser in objGroup.Members : RET_73 = RET_73 & objUser.Name & NEWLINE_CHARACTER: Next
	IF objGroup.Name = "Guests" THEN For Each objUser in objGroup.Members : RET_74 = RET_74 & objUser.Name & NEWLINE_CHARACTER: Next
	IF objGroup.Name = "Replicator" THEN For Each objUser in objGroup.Members : RET_75 = RET_75 & objUser.Name & NEWLINE_CHARACTER: Next
	IF objGroup.Name = "Users" THEN For Each objUser in objGroup.Members : RET_76 = RET_76 & objUser.Name & NEWLINE_CHARACTER: Next
Next

	IF RET_71="" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = MANUAL_CHARACTER()
	WRITE_OUTPUT_5 "7.1", "Administrators", REQ_71, RET_71, strCompare

	IF RET_72="" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = MANUAL_CHARACTER()
	WRITE_OUTPUT_5 "7.2", "Backup Operators", REQ_72, RET_72, strCompare
	
	IF RET_73="" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 "7.3", "Power Users", REQ_73, RET_73, strCompare
	
	IF InStr("CUSTOMGUEST<br/>",RET_74) <> 0 THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 "7.4", "Guests", REQ_74, RET_74, strCompare
	
	IF RET_75="" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = MANUAL_CHARACTER()
	WRITE_OUTPUT_5 "7.5", "Replicator", REQ_75, RET_75, strCompare
	
	IF RET_76="" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = MANUAL_CHARACTER()
	WRITE_OUTPUT_5 "7.6", "Users", REQ_76, RET_76, strCompare
WRITE_OUTPUT_SECTION_END "SECTION 7"

'================================================================================
'  SECTION 8
'================================================================================
strCompare = NA_CHARACTER()
WRITE_OUTPUT_SECTION_START "SECTION 8 - System Group Accounts", ""
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"
	
	IF TASK08_RETURN(0)="SeDenyNetworkLogonRight<br/>" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 "8.1", "Anonymous Logon", "-", TASK08_RETURN(0), strCompare
	
	IF TASK08_RETURN(1)="SeNetworkLogonRight<br/>SeChangeNotifyPrivilege<br/>" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = MANUAL_CHARACTER()
	WRITE_OUTPUT_5 "8.2", "Authenticated Users", "-", TASK08_RETURN(1), strCompare
	
	IF TASK08_RETURN(2)="" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 "8.3", "DIALUP", "-", TASK08_RETURN(2), strCompare
	
	IF TASK08_RETURN(3)="SeImpersonatePrivilege<br/>SeCreateGlobalPrivilege<br/>" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 "8.4", "SERVICE", "-", TASK08_RETURN(3), strCompare
	'
	IF InStr(TASK08_RETURN(4),"Domain Controller") <> 0 THEN strCompare = MANUAL_CHARACTER() & "(DC FOUND!)" ELSE strCompare = NA_CHARACTER() & " - NOT A DC (" & strDomainRole & ")"
	WRITE_OUTPUT_5 "8.5", "SELF", "-", TASK08_RETURN(4), strCompare
	
	IF TASK08_RETURN(5)="" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 "8.6", "INTERACTIVE", "-", TASK08_RETURN(5), strCompare
	
	IF TASK08_RETURN(6)="" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 "8.7", "Everyone", "-", TASK08_RETURN(6), strCompare
	
	IF TASK08_RETURN(7)="" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 "8.8", "TERMINAL SERVER USER", "-", TASK08_RETURN(7), strCompare
WRITE_OUTPUT_SECTION_END "SECTION 8"


'================================================================================
'  SECTION 9
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 9 - Default User Accounts", ""
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"
	
	REQ_91 = "CUSTOMADMIN"
	RES_91 = task_44
	IF InStr(REQ_91,RES_91) <> 0 THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 "9.1", "Administrator", REQ_91, RES_91, strCompare
	
	REQ_92 = "CUSTOMGUEST"
	RES_92 = task_45
	IF InStr(REQ_92,RES_92) <> 0 THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 "9.2", "Guest", REQ_92, RES_92, strCompare
	
	strCompare = NA_CHARACTER()
	
	
	
	strCurrentValue12_1 = ""
	strCompare12_1 = WARNING_CHARACTER() & "NOT CHECKED?"
	
	Set colGroups = GetObject("WinNT://.")
	colGroups.Filter = Array("group")
	For Each objGroup In colGroups
		IF objGroup.Name = "Remote Desktop Users" THEN
			For Each objUser in objGroup.Members
				strCurrentValue12_1 = strCurrentValue12_1 & objUser.Name & NEWLINE_CHARACTER
			Next
			
			IF strCurrentValue12_1="" THEN strCompare12_1 = CONFIRM_CHARACTER() ELSE strCompare12_1 = ERROR_CHARACTER()
		END IF
	Next
	
	WRITE_OUTPUT_5 "9.3", "Remote Desktop Users", "-", strCurrentValue12_1, strCompare12_1
WRITE_OUTPUT_SECTION_END "SECTION 9"

'================================================================================
'  SECTION 10
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 10 - Access to Volumes / Shares", "Volume & Folder security, check their respective ACL settings"
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"


	'ANOTEXISTED = Array("NOT EXISTED")
	'AREQ_TYPE_2 = Array("0:BUILTIN\Administrators:2032127", "0:NT AUTHORITY\SYSTEM:2032127")
	'AREQ_TYPE_2C= Array("0:BUILTIN\Administrators:2032127", "0:NT AUTHORITY\SYSTEM:2032127", "0:\CREATOR OWNER:2032127")
	'AREQ_TYPE_3 = Array("0:BUILTIN\Administrators:2032127", "0:NT AUTHORITY\SYSTEM:2032127", "0:NT AUTHORITY\Authenticated Users:1179817")
	'AREQ_TYPE_4 = Array("0:BUILTIN\Administrators:2032127", "0:NT AUTHORITY\SYSTEM:2032127", "0:NT AUTHORITY\Authenticated Users:1179817", "0:\CREATOR OWNER:2032127")
	'AREQ_TYPE_5 = Array("0:BUILTIN\Administrators:2032127", "0:NT AUTHORITY\SYSTEM:2032127", "0:NT AUTHORITY\Authenticated Users:1245631", "0:\CREATOR OWNER:2032127", "0:BUILTIN\Replicator:1245631")


	'REQ_TYPE_2 = "0:BUILTIN\Administrators:2032127;0:NT AUTHORITY\SYSTEM:2032127;"
	'REQ_TYPE_2_REV= "0:NT AUTHORITY\SYSTEM:2032127;0:BUILTIN\Administrators:2032127;"
	'REQ_TYPE_2_CO = "0:BUILTIN\Administrators:2032127;0:\CREATOR OWNER:2032127;0:NT AUTHORITY\SYSTEM:2032127;"
	'REQ_TYPE_3 = "0:BUILTIN\Administrators:2032127;0:NT AUTHORITY\Authenticated Users:1179817;0:NT AUTHORITY\SYSTEM:2032127;"
	'REQ_TYPE_4 = "0:BUILTIN\Administrators:2032127;0:NT AUTHORITY\Authenticated Users:1179817;0:\CREATOR OWNER:2032127;0:NT AUTHORITY\SYSTEM:2032127;"
	'REQ_TYPE_5 = "0:BUILTIN\Administrators:2032127;0:NT AUTHORITY\Authenticated Users:1245631;0:\CREATOR OWNER:2032127;0:BUILTIN\Replicator:1245631;0:NT AUTHORITY\SYSTEM:2032127;"

	Set objShell = CreateObject("WScript.Shell")
	Set objEnv = objShell.Environment("Process")

	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\"
	CACLS_REPORT "10.1", AREQ_TYPE_4, AUDIT_FILENAME

	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colDisks = objWMIService.ExecQuery("SELECT * from Win32_LogicalDisk WHERE DriveType=3 AND NOT DeviceID='C:'")
	Dim count : count = 0
	For Each objDisk in colDisks
		count = count + 1
		AUDIT_FILENAME = objDisk.DeviceID & "\"
		CACLS_REPORT "10.2", AREQ_TYPE_4, AUDIT_FILENAME
	Next
	IF count = 0 THEN WRITE_OUTPUT_5 "10.2", "Additional Drive", "-", "NO OTHER EXTERNAL DRIVE FOUND", NA_CHARACTER()
	'wscript.quit

	WRITE_OUTPUT_5 "10.3", "SUBFOLDER OF A VOLUME", "-", "-", NA_CHARACTER()
	WRITE_OUTPUT_5 "10.4", "WHEN SHARE CREATED", "-", "-", NA_CHARACTER()
	WRITE_OUTPUT_5 "10.5", "INHERITANCE OF SHARE", "-", "-", NA_CHARACTER()
	
	
	Function ShowFolderList(folderspec)
		Dim fso, f, Fldr, fc, s
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.GetFolder(folderspec)
		Set fc = f.Files
		
		Set Fldrs = f.SubFolders  
	  
	  
	  
		LIST_OF_10_6_SUCCESS = CONFIRM_CHARACTER() & "<b>SUCCESS:</b>" & NEWLINE_CHARACTER()
		LIST_OF_10_6_ERROR = ERROR_CHARACTER() & "<b>NOT FULFILLED:</b>" & NEWLINE_CHARACTER()
				
		TOTAL_10_6 = 0
		For Each Fldr In Fldrs  
			IF InStr(Fldr.name,"$NtUninstall") <> 0 THEN 
				TOTAL_10_6 = TOTAL_10_6 + 1
				'WScript.Echo Fldr
				AUDIT_FILENAME = Fldr
				
				
				IF CACLS_REPORT_SILENCE("10.6", AREQ_TYPE_2, AUDIT_FILENAME) THEN
					LIST_OF_10_6_SUCCESS = LIST_OF_10_6_SUCCESS & Fldr.name & NEWLINE_CHARACTER
				ELSE
					LIST_OF_10_6_ERROR = LIST_OF_10_6_ERROR & Fldr.name & NEWLINE_CHARACTER
				END IF
				'CACLS_REPORT "10.6", AREQ_TYPE_2, AUDIT_FILENAME
				'strRequireValue = REQ_TYPE_2
				'strCurrentValue = getFilePermissions(AUDIT_FILENAME, strTEMPValue)
				'IF InStr(strTEMPValue,strRequireValue) <> 0 OR InStr(strTEMPValue, REQ_TYPE_2_REV) THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
				'IF strCompare = ERROR_CHARACTER() AND strRequireValue = REQ_TYPE_2 AND InStr(strTEMPValue,REQ_TYPE_2_CO) <> 0 THEN strCompare = CONFIRM_CHARACTER() & " Creator Owner: FULL"
				'IF strCompare = ERROR_CHARACTER() THEN strCompare = MANUAL_CHARACTER()
				'WRITE_OUTPUT_5 "10.6", AUDIT_FILENAME, getHTMLFilePerm(strRequireValue), strCurrentValue, strCompare
			END IF
		Next
	
		IF LIST_OF_10_6_ERROR = ERROR_CHARACTER() & "<b>NOT FULFILLED:</b>" & NEWLINE_CHARACTER() THEN LIST_OF_10_6_ERROR = CONFIRM_CHARACTER()
		WRITE_OUTPUT_5 "10.6", folderspec & "\$NtUninstall...", getHTMLFilePerm(Join(AREQ_TYPE_2, ";")), getHTMLFilePerm(Join(AREQ_TYPE_2, ";")), LIST_OF_10_6_ERROR & " TOTAL: " & TOTAL_10_6
		
		ShowFolderList = s
	End Function

	IF SKIP_10_6 = TRUE THEN
		WRITE_OUTPUT_5 "10.6", AUDIT_FILENAME, "", "", WARNING_CHARACTER() & " - FORCE SKIPPED (config.ini)"
	ELSE
		wscript.stderr.write ""
		ShowFolderList(objEnv.Item("SystemRoot"))
		'AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\$NtUninstall*"
		'WRITE_OUTPUT_5 "10.6", AUDIT_FILENAME, "", getFilePermissions(AUDIT_FILENAME, strTEMPValue), strCompare
		'strCompare = NA_CHARACTER()
		'WRITE_OUTPUT_5 "10.6", AUDIT_FILENAME, "", "", UNDERCONSTRUCTION_CHARACTER()
	END IF


	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\Documents and Settings"
	CACLS_REPORT "10.7", AREQ_TYPE_4, AUDIT_FILENAME
	
	
	AREQ_TYPE_SP = Array( _
	"0:" & strHostname & "\" & task_44 & ":2032127", _
	"0:NT AUTHORITY\SYSTEM:2032127", _
	"0:BUILTIN\Administrators:2032127", _
	"0:" & strHostname & "\" & task_44 & ":268435456", _
	"0:NT AUTHORITY\SYSTEM:268435456", _
	"0:BUILTIN\Administrators:268435456" _
	)

	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\Documents and Settings\Administrator"
	IF CACLS_REPORT_SILENCE("10.8", AREQ_TYPE_2, AUDIT_FILENAME) THEN
		CACLS_REPORT "10.8", AREQ_TYPE_2, AUDIT_FILENAME
	ELSE
		CACLS_REPORT "10.8", AREQ_TYPE_SP, AUDIT_FILENAME
	END IF
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\Documents and Settings\All Users"
	CACLS_REPORT "10.9", AREQ_TYPE_3, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\Documents and Settings\Default User"
	CACLS_REPORT "10.10", AREQ_TYPE_3, AUDIT_FILENAME
	
	strAuditNo = "10.11"
	strKeyName = "POSIX FILE REMOVE"
	strRequire = "NOT FOUND"
	IF objFSO.FileExists("C:\Windows\System32\Os2.exe") OR _
	objFSO.FileExists("C:\Windows\System32\Os2ss.exe") OR _
	objFSO.FileExists("C:\Windows\System32\Os2srv.exe") OR _
	objFSO.FileExists("C:\Windows\System32\Psxss.exe") OR _
	objFSO.FileExists("C:\Windows\System32\Posix.exe") OR _
	objFSO.FileExists("C:\Windows\System32\Psxdll.dll") Then value = "ERROR! FILE(S) EXIST." ELSE value = "NOT FOUND"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, getHTMLFilePerm(strRequireValue), value, strCompare

	strCompare = ""

	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32"
	CACLS_REPORT "10.12", AREQ_TYPE_4, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\config"
	CACLS_REPORT "10.13", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\Ntbackup.exe"
	CACLS_REPORT "10.14", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\rcp.exe"
	CACLS_REPORT "10.15", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\Rdisk.exe"
	CACLS_REPORT "10.16", ANOTEXISTED, AUDIT_FILENAME

	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\Regedt32.exe"
	CACLS_REPORT "10.17", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\Regedt.cnt"
	CACLS_REPORT "10.17", ANOTEXISTED, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\Regedt32.hlp"
	CACLS_REPORT "10.17", ANOTEXISTED, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\repl\export"
	CACLS_REPORT "10.18", ANOTEXISTED, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\repl\import"
	CACLS_REPORT "10.19", ANOTEXISTED, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\rexec.exe"
	CACLS_REPORT "10.20", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\rsh.exe"
	CACLS_REPORT "10.21", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\system32\spool\Printers"
	CACLS_REPORT "10.22", AREQ_TYPE_5, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\"
	CACLS_REPORT "10.23", AREQ_TYPE_4, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\autoexec.bat"
	CACLS_REPORT "10.24", AREQ_TYPE_3, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\config.sys"
	CACLS_REPORT "10.25", AREQ_TYPE_3, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\io.sys"
	CACLS_REPORT "10.26", AREQ_TYPE_3, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\msdos.sys"
	CACLS_REPORT "10.27", AREQ_TYPE_3, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\ntdetect.com"
	CACLS_REPORT "10.28", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\ntldr"
	CACLS_REPORT "10.29", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\NTReskit"
	CACLS_REPORT "10.30", ANOTEXISTED, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemDrive") & "\Program Files"
	CACLS_REPORT "10.31", AREQ_TYPE_4, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot")
	CACLS_REPORT "10.32", AREQ_TYPE_4, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\$NtServicePackUninstall$"
	CACLS_REPORT "10.33", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\Cookies"
	CACLS_REPORT "10.34", ANOTEXISTED, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\Help"
	CACLS_REPORT "10.35", AREQ_TYPE_4, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\History"
	CACLS_REPORT "10.36", ANOTEXISTED, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\regedit.exe"
	CACLS_REPORT "10.37", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\repair"
	CACLS_REPORT "10.38", AREQ_TYPE_2, AUDIT_FILENAME
	
	AUDIT_FILENAME = objEnv.Item("SystemRoot") & "\Security"
	CACLS_REPORT "10.39", AREQ_TYPE_2, AUDIT_FILENAME
	
	'Set objShell = Nothing
	'Set objEnv = Nothing
WRITE_OUTPUT_SECTION_END "SECTION 10"



'================================================================================
'  SECTION 12
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 12 - Configuring RDP", "Skip this section if RDP is not enabled."
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"

	
	'strCurrentValue12_1 = ""
	strCurrentValue12_2 = ""
	strCurrentValue12_3 = ""
	
	'strCompare12_1 = WARNING_CHARACTER() & "NOT CHECKED?"
	strCompare12_2 = NA_CHARACTER()
	strCompare12_3 = NA_CHARACTER()
	
	Set colGroups = GetObject("WinNT://.")
	colGroups.Filter = Array("group")
	For Each objGroup In colGroups
		strCurrent = ""
		IF objGroup.Name = "TSUSERS" THEN
			strCurrentValue12_2 = "TSUSERS"
			strCompare12_2 = CONFIRM_CHARACTER()
			strCompare12_3 = MANUAL_CHARACTER()
			For Each objUser in objGroup.Members
				strCurrentValue12_3 = strCurrentValue12_3 & objUser.Name & NEWLINE_CHARACTER
			Next
			Exit For
		END IF
		
		'IF objGroup.Name = "Remote Desktop Users" THEN
		'	For Each objUser in objGroup.Members
		'		strCurrentValue12_1 = strCurrentValue12_1 & objUser.Name & NEWLINE_CHARACTER
		'	Next
			
		'	IF strCurrentValue12_1="" THEN strCompare12_1 = CONFIRM_CHARACTER() ELSE strCompare12_1 = ERROR_CHARACTER()
		'END IF
	Next

	WRITE_OUTPUT_5 "12.1", "Disable TSInternetUsers", "Disabled [no user]", strCurrentValue12_1, strCompare12_1
	WRITE_OUTPUT_5 "12.2", "Create User Group TSUSERS", "TSUSERS", strCurrentValue12_2, strCompare12_2
	WRITE_OUTPUT_5 "12.3", "NO GUEST IN TSUSERS Group", "ONLY AUTHORISED USERS", strCurrentValue12_3, strCompare12_3
	
	
	
	strKeyPath = "SYSTEM\ControlSet001\Control\Terminal Server\WinStations\RDP-Tcp"
	strKeyName = Array("MaxDisconnectionTime", "MaxConnectionTime", "MaxIdleTime", "fInheritMaxSessionTime", "fInheritMaxDisconnectionTime", "fInheritMaxIdleTime")
	strRequire = Array(60000, 0, 0, 0, 0, 0)
	strCompare = CONFIRM_CHARACTER()
	FOR x=0 TO ubound(strKeyName)
		IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName(x), strTEMPValue) <> 0 Then strCompare = ERROR_CHARACTER() & " - " & strKeyName(x) & "NOT EXISTED" : Exit For
		IF NOT (strTEMPValue = strRequire(x)) Then strCompare = ERROR_CHARACTER() & " - " & strKeyName(x) & ": " & strTEMPValue : Exit For
	NEXT
	WRITE_OUTPUT_5 "12.4", "Timeout Settings", "Disconnected Session: 1min<br/>Active Session: Never<br/>Idle Session: Never", "-", strCompare
	
	IF SILENCE = TRUE THEN strCurrentValue = "[SILENCE - SKIPPED]" ELSE strCurrentValue = InputBox("12.5 : Terminal Services Permission", "MANUAL 12.5", "Administrators, System")
	IF strCurrentValue="Administrators" OR _
	   strCurrentValue="Administrators, System" OR _
	   strCurrentValue="Administrators, System, TSUSERS" _
	THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = MANUAL_CHARACTER()
	WRITE_OUTPUT_5 "12.5", "TS Permission", "Administrators, System, TSUSERS", strCurrentValue, strCompare & " (ENTERED MANUALLY)"
	
	strCurrentValue = Replace(ReplaceSID(TASK33_QUERY(5)),",",NEWLINE_CHARACTER)
	IF strCurrentValue="Administrators" OR _
	   strCurrentValue="Administrators, System" OR _
	   strCurrentValue="Administrators, System, TSUSERS" _
	THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = MANUAL_CHARACTER()
	IF strCurrentValue="Administrators" THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = MANUAL_CHARACTER()
	WRITE_OUTPUT_5 "12.6", TASK33(5), "Administrators<br/>System<br/>TSUSERS", strCurrentValue, strCompare
	
WRITE_OUTPUT_SECTION_END "SECTION 12"


'================================================================================
'  SECTION 13
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 13 - Miscellaneous", ""
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"

	strAuditNo = "13.1"
	strKeyPath = "SYSTEM\CurrentControlSet\Control\FileSystem"
	strKeyName = "NtfsDisable8dot3NameCreation"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "13.2"
	strKeyName = "timeout"
	strRequire = "0"
	value = ReadInI("C:\boot.ini", "boot loader", strKeyName)
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "13.3"
	strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager"
	strKeyName = "ProtectionMode"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

	strAuditNo = "13.4"
	strKeyPath = "SYSTEM\CurrentControlSet\Control\Lsa"
	strKeyName = "SecureBoot"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare

WRITE_OUTPUT_SECTION_END "SECTION 13"



'================================================================================
'  SECTION 17
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 17 - Hardening IIS", ""
WRITE_TH "No", "Description", "Required Value", "Current Value", "Remark"

	strAuditNo = "17.1"
	strKeyPath = "SYSTEM\CurrentControlSet\Services\W3SVC\Parameters"
	strKeyName = "DisableWebDAV"
	strRequire = 1
	IF oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strKeyName, value) <> 0 Then value = "NOT EXISTED"
	IF value="NOT EXISTED" THEN
		strCompare = NA_CHARACTER()
	ELSE
		IF strRequire=value THEN strCompare = CONFIRM_CHARACTER() ELSE strCompare = ERROR_CHARACTER()
	END IF
	WRITE_OUTPUT_5 strAuditNo, strKeyName, strRequire, value, strCompare
WRITE_OUTPUT_SECTION_END "SECTION 17"





'================================================================================
'  SECTION LOCAL GROUPS
'================================================================================
WRITE_OUTPUT_SECTION_START "SECTION 18 - Local Groups", "List of local groups and their respective assigned users"
WRITE_TH "No", "Group", "User", "&nbsp;", "&nbsp;"
	count = 0
	Set colGroups = GetObject("WinNT://.")
	colGroups.Filter = Array("group")
	For Each objGroup In colGroups
		count = count + 1
		strCurrent = ""
		For Each objUser in objGroup.Members
			strCurrent = strCurrent & objUser.Name & NEWLINE_CHARACTER
		Next
		WRITE_OUTPUT_5 "18." & count, objGroup.Name, strCurrent, "", ""
	Next
WRITE_OUTPUT_SECTION_END "LOCAL GROUPS"


IF HTML_MODE = TRUE THEN
	wscript.stdout.write "<p class=""copyright"">&copy;2009 YouQi.</p></div></body></html>"
END IF