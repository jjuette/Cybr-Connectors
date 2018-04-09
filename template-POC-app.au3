
Global $TargetUsername
Global $TargetPassword
Global $ConnectionClientPID = 0
Global $sleep = 1000

;POC - Start
Global Const $CLIENT_EXECUTABLE  = "executable path and name" ; CHANGE_ME
;POC - END


Local $ConnectionClientPID = Run($CLIENT_EXECUTABLE, "", @SW_MAXIMIZE)
if ($ConnectionClientPID == 0) Then
Error(StringFormat("Failed to execute process [%s]", $CLIENT_EXECUTABLE, @error))
EndIf



;POC - Start
Local $TITLE = "<Title>"

WinWaitActive($TITLE)

; Controls
Local $CONTROLID_LOGIN 			= "[CLASSNN:<ClassName>]"
Local $CONTROLID_PASSWORD 		= "[CLASSNN:<ClassName>]"
Local $CONTROLID_BUTTON 		= "[CLASSNN:<ClassName>]"


ControlSetText($TITLE, "", $CONTROLID_LOGIN , $TargetUsername)
ControlSetText($TITLE, "", $CONTROLID_PASSWORD , $TargetPassword)
ControlClick($TITLE, "",  $CONTROLID_BUTTON )

;POC - END

