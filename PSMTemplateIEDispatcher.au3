#AutoIt3Wrapper_UseX64=n
Opt("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", 3) ; EXACT_MATCH!

;============================================================
;             PSM PrivateArk Client Dispatcher
;             ------------------------------
; Description : PSM Dispatcher for the PrivateArk Client
; Created : October 2015
; Cyber-Ark Software Ltd.
; Developed and compiled on AutoIt 3.3.14.1
;============================================================

#include "PSMGenericClientWrapper.au3"
#include <IE.au3>

;================================
; Consts & Globals
;================================

Global Const $WEBSITE_NAME 	= "POC Connector"
Global Const $MESSAGE_TITLE	= "PSM - POC Connector"

;Parameters
Global $g_pTemplateAddress = "%s/PasswordVault/auth/cyberark"

; The Connection Parametrs will be saved here
; Parameters that are not included in the Password Object will use these Defaults
Global $g_pTargetUsername	= ""
Global $g_pTargetPassword	= ""
Global $g_pTargetAddress	= ""

Global $g_pIELoadWait_Delay = 900 ; Change if needed
Global $g_pIELoadWait_Timeout = 30000 ; Change if needed


;================================
; Consts & Globals - DO NOT CHANGE
;================================

Global Const $ERROR_MESSAGE_TITLE = "PSM " & $WEBSITE_NAME & " Dispatcher error message"
Global Const $LOG_MESSAGE_PREFIX  = $WEBSITE_NAME & " Dispatcher - "

Global $g_ConnectionClientPID = 0
Global $g_IEInstance

;Internal Variables
Global $g_ErrorMessageTime = 15000

; Start the Code
Exit Main()

Func LoginProcess()
	LogWrite("Entered LoginProcess()")

	If (IsCertificateErrorPage($g_IEInstance)) Then
		ContinueOnCertificateError($g_IEInstance)
    EndIf

    _IEAction($g_IEInstance, "visible")
	_IELoadWait($g_IEInstance,$g_pIELoadWait_Delay,$g_pIELoadWait_Timeout)
	

	; POC GOES HERE
	;
	;
	;
	;
	;
	;
	;
	;
	;;;;;;;;;;;;;;;;;;
	

    LogWrite("finished LoginProcess() successfully")
    UnblockAllBlockProhibited()
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: FetchSessionProperties
; Description ...: Fetches properties required for the session from the PSM
; Parameters ....: None
; Return values .: None
; ===============================================================================================================================
Func FetchSessionProperties()
	; Get the Session User Name
	If (PSMGenericClient_GetSessionProperty("Username", $g_pTargetUsername) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf
	
	; Get the Session Password
	If (PSMGenericClient_GetSessionProperty("Password", $g_pTargetPassword) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	; Get the Session Address
	If (PSMGenericClient_GetSessionProperty("Address", $g_pTargetAddress) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
    Else
        LogWrite("PVWAAddress="&$g_pTargetAddress)
    EndIf
EndFunc

;*=*=*=*=*=*=*=*=*=*=*=*=*=
; DO NOT CHANGE FROM HERE
;*=*=*=*=*=*=*=*=*=*=*=*=*=

;=======================================
; Main
;=======================================
Func Main()
	; Init PSM Dispatcher utils wrapper
    MessageUserOn($MESSAGE_TITLE, "The PSM is about to log you on automatically which may take several seconds...")

	If (PSMGenericClient_Init() <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	LogWrite("successfully initialized Dispatcher Utils Wrapper")
	LogWrite("mapping local drives")

	If (PSMGenericClient_MapTSDrives() <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
	EndIf

	; Get the dispatcher parameters
	FetchSessionProperties()

	MessageUserOn($MESSAGE_TITLE, "Starting " & $WEBSITE_NAME & "...")

	LogWrite("starting client application")
    ;Open IE. OpenIE will terminate the process If IE wasn't loaded properly. No need to check PID here.
	$g_ConnectionClientPID = OpenIE()

    ;Send PID to PSM as early as possible so recording/monitoring can begin
	LogWrite("sending PID to PSM")
	If (PSMGenericClient_SendPID($g_ConnectionClientPID) <> $PSM_ERROR_SUCCESS) Then
		Error(PSMGenericClient_PSMGetLastErrorString())
    EndIf

	;Handle Login
	MessageUserOff()
    LoginProcess()

	; Terminate PSM Dispatcher utils wrapper
	LogWrite("Terminating Dispatcher Utils Wrapper")
	PSMGenericClient_Term()

	Return $PSM_ERROR_SUCCESS
EndFunc

;==================================
; Functions
;==================================

; #FUNCTION# ====================================================================================================================
; Name...........: Error
; Description ...: An exception handler - displays an error message and terminates the dispatcher
; Parameters ....: $ErrorMessage - Error message to display
; 				   $Code 		 - [Optional] Exit error code
; ===============================================================================================================================
Func Error($ErrorMessage, $Code = -1)

   ; If the dispatcher utils DLL was already initialized, write an error log message and terminate the wrapper
   If (PSMGenericClient_IsInitialized()) Then
		LogWrite($ErrorMessage, $LOG_LEVEL_ERROR)
		PSMGenericClient_Term()
   EndIf

   MessageUserOn("ERROR - PROCESS IS SHUTTING DOWN", $ErrorMessage)
	sleep($g_ErrorMessageTime)
	; If the connection component was already invoked, terminate it
	If ($g_ConnectionClientPID <> 0) Then
		ProcessClose($g_ConnectionClientPID)
		$g_ConnectionClientPID = 0
	EndIf
	Exit $Code

EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: LogWrite
; Description ...: Write a PSMWinSCPDispatcher log message to standard PSM log file
; Parameters ....: $sMessage - [IN] The message to write
;                  $LogLevel - [Optional] [IN] Defined If the message should be handled as an error message or as a trace messge
; Return values .: $PSM_ERROR_SUCCESS - Success, otherwise error - Use PSMGenericClient_PSMGetLastErrorString for details.
; ===============================================================================================================================
Func LogWrite($sMessage, $LogLevel = $LOG_LEVEL_TRACE)
	Return PSMGenericClient_LogWrite($LOG_MESSAGE_PREFIX & $sMessage, $LogLevel)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: AssertErrorLevel
; Description ...: Checks If error level is <> 0. If so, write to log and call error.
; Parameters ....: $error_code - the error code from last function call (@error)
;                  $message - Message to show to user as well as write to log
;				   $code - exit code (default -1)
; Return values .: None
; ===============================================================================================================================
Func AssertErrorLevel($error_code, $message, $code = -1)
   ;Unblock input so user can exit
	If ($error_code <> 0) Then
		LogWrite(StringFormat("AssertErrorLevel - %s :: @error = %d", $message, $error_code), $LOG_LEVEL_ERROR)
		Error($message, $code)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: MessageUserOn
; Description ...: Writes a message to the user, and keeps it indefinitely (until function call to MessageUserOff)
; Parameters ....: $msgTitle - Title of the message
;                  $msgBody - Body of the message
; Return values .: none
; ===============================================================================================================================
Func MessageUserOn(Const ByRef $msgTitle, Const ByRef $msgBody)
	SplashOff()
    SplashTextOn ($msgTitle, $msgBody, -1, 54, -1, -1, 0, "Tahoma", 9, -1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: MessageUserOff
; Description ...: See SplashOff()
; Parameters ....:
;
; Return values .: none
; ===============================================================================================================================
Func MessageUserOff()
    SplashOff()
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: BlockAllInput
; Description ...: Blocks all input (mouse & keyboard). Use when login process runs and visible, so user can't
;                  manipulate the process
; Parameters ....:
; Return values .: none
; ===============================================================================================================================
Func BlockAllInput()
    LogWrite("Blocking Input")
	;Block all input - mouse and keyboard
	If IsDeclared("s_KeyboardKeys_Buffer") <> 0 Then
		_BlockInputEx(1)
		AssertErrorLevel(@error, StringFormat("Could not block all input. Aborting... @error: %d", @error))
	Else
		BlockInput(1)
		AssertErrorLevel(@error, StringFormat("Could not block all input. Aborting... @error: %d", @error))
	EndIf

EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: UnblockAllBlockProhibited
; Description ...: Allows all input from the user, except for prohibited keys (such as F11).
; Parameters ....:
;
; Return values .: none
; ===============================================================================================================================
Func UnblockAllBlockProhibited()
	If IsDeclared("s_KeyboardKeys_Buffer") <> 0 Then
		_BlockInputEx(0)
		_BlockInputEx(3, "", "{F11}|{Ctrl}") ;Ctrl - also +C? +V?...
	Else
		BlockInput(0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: OpenIE
; Description ...: Creates Internet explorer and navigates to the address
; Parameters ....: None
; Return values .: The Program Process ID (For PSM Logging)
; ===============================================================================================================================
Func OpenIE()
	; Translate the Web Site URL Address with the requiered device
	;dim $_TargetURL = StringFormat($g_pTemplateAddress, $g_pTargetAddress)
    dim $_TargetURL = $g_pTargetAddress

	; Fix issue in Win2012 - Turn off Protection Mode
	FixProtectionMode()

	; Fix issue with Cookies
	FixCookies()

	;Open a new IE instance which is invisible and navigate to $_TargetURL, then wait for IE to finish loading.
	$g_IEInstance = _IECreate($_TargetURL, Default, 0)

	AssertErrorLevel(@error, "Call to _IECreate failed", -1002)
    _IELoadWait($g_IEInstance,$g_pIELoadWait_Delay,$g_pIELoadWait_Timeout)
	;HardenIE()
	local $hndl = _IEPropertyGet($g_IEInstance, "hwnd")
	return HWindowToPID($hndl)
 EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: HWindowToPID
; Description....: Returns the PID that owns the given window HWND.
; Parameters.....: $hWindow - HWND of the window
; Return values..: PID of the process that owns the window. If PID = 0 this function calls Error thus terminating the
;                  autoit process
; ===============================================================================================================================
Func HWindowToPID($hWindow)
	Local $pPID
	Local $dwPID
	Local $result

	$pPID = DllStructCreate("DWORD")

	; DWORD GetWindowThreadProcessId(HWND hWnd, LPDWORD lpdwProcessId)
	; Parameters:
	;    hWnd 			[IN]  HWND of the window
	;    lpdwProcessId  [OUT] Process ID
	; Return value: Thread ID (DWORD)
	$result = DllCall("user32.dll", "DWORD", "GetWindowThreadProcessId", "hwnd", $hWindow, "ptr", DllStructGetPtr($pPID))

	If (@error <> 0) Then
		$pPID = 0
		Error(StringFormat("Failed to get IE PID (Extra details: %d)", @error),-20)
	EndIf
	$dwPID = DllStructGetData($pPID, 1)
	$pPID = 0

	return $dwPID
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: HardenIE
; Description ...: This function makes the IE fullscreen and disables toolbar, menubar and address bar. Note that If you do not use _ExBlockInput,
;                  user can press F11 key and get to the address bar. To solve this, use BlockAllInput(), AllowAllInputButF11()
; Parameters ....: None
; Return values .: None
; ===============================================================================================================================
Func HardenIE()
   SetIEProperty("theatermode",0)
   SetIEProperty("toolbar", 0)
   SetIEProperty("menubar", 0)
   SetIEProperty("addressbar", 0)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: SetIEProperty
; Description....: Internal utility function
; ===============================================================================================================================
Func SetIEProperty($sProperty, $nValue)
   dim $g_IEInstance
	If(_IEPropertySet($g_IEInstance, $sProperty, $nValue) <> $_IEStatus_Success) Then
		Error(StringFormat("WebFormDispatcher: Failed to set IE property [%s] to [%s] (Extra details: %d, %d)", $sProperty, $nValue, @error, @extended))
	EndIf
 EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: FixCookies
; Description....: Sets an IE Zones Regestry Value to fix saving Cookies by default
; Parameters.....: NONE
; Return values..: NONE
; ===============================================================================================================================
Func FixCookies()
	For $i = 1 to 4
		SetIEZoneRegDWORDValue($i, "1601", "0")
		SetIEZoneRegDWORDValue($i, "1A10", "1")
	Next
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: FixProtectionMode
; Description....: Sets an IE Zones Regestry Value to disable Protection Mode
; Parameters.....: NONE
; Return values..: NONE
; ===============================================================================================================================
Func FixProtectionMode()
	For $i = 1 to 4
		SetIEZoneRegDWORDValue($i, "2500", "0")
	Next
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: SetIEZoneRegDWORDValue
; Description....: Sets an IE Zones Regestry Value
; Parameters.....: $nZone - Zone Number
;				   $sName - Name
;				   $sValue - Value
; Return values..: NONE
; ===============================================================================================================================
Func SetIEZoneRegDWORDValue($nZone, $sName, $sValue)
	RegWrite(StringFormat("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\%s", $nZone) , $sName, "REG_DWORD", $sValue)
	AssertErrorLevel(@error, "Call RegWrite", @error)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: IsCertificateErrorPage
; Description ...: This function checks If the IE object is on a website security certificate error page.
; Parameters ....: $IEObject - the IE instance to check
; Return values .: True If on certification error page, False otherwise
; ===============================================================================================================================
Func IsCertificateErrorPage($IEObject)
	local $returnVal = False
	If StringInStr(_IEBodyReadText ($IEObject), "security certificate") Then
		$returnVal = True
	EndIf
	return $returnVal
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: ContinueOnCertificationError
; Description ...: This function clicks on "continue to this website (not recommended)" link on certification error page.
; Parameters ....: $IEObject - the IE instance
; Return values .: none
; ===============================================================================================================================
Func ContinueOnCertificateError($IEObject)
	_IELinkClickByText ($IEObject, "Continue to this website (not recommended).",0,0)
EndFunc