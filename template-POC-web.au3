#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.12.0
 Author:         myName

 Script Function:
	POC AutoIt script.

#ce ----------------------------------------------------------------------------
#include <IE.au3>

Local $g_pTargetUsername = "username"
Local $g_pTargetPassword = "password"
Local $Address = "https://login-page-here"
; Script Start - Add your code below here

 Local $g_IEInstance 			= _IECreate($Address)

 
 
;POC Code starts here
 
Local $UsernameId = "<userID>"
Local $PasswordId = "<passID>"
Local $LoginBtnId = "<loginID>"

; by ID
 Local $oUserName 		= _IEGetObjByID($g_IEInstance, $UsernameId)
 Local $oPassword 		= _IEGetObjByID($g_IEInstance, $PasswordId)
 Local $LoginBtn 		= _IEGetObjByID($g_IEInstance, $LoginBtnId)
 
; by Name
; Local $oUserName 		= _IEGetObjByName($g_IEInstance, $UsernameId)
; Local $oPassword 		= _IEGetObjByName($g_IEInstance, $PasswordId)
; Local $LoginBtn 		= _IEGetObjByName($g_IEInstance, $LoginBtnId)

_IEFormElementSetValue($oUserName, $g_pTargetUsername)
_IEFormElementSetValue($oPassword, $g_pTargetPassword)

; By click
_IEAction($LoginBtn,'click')

; By ENTER
;sleep(1000)
;send("{ENTER}")