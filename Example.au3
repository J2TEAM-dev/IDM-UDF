#NoTrayIcon
#include <IDM.au3>

Opt('MustDeclareVars', 1)

example()

Func example()
	; Make sure we can use IDM
	If $IDM_Init_Status Then
		Local $sURL, $sLocalPath, $sLocalFileName, $sFlags
		$sURL = 'http://www.internetdownloadmanager.com/trans_kit.zip'
		$sLocalPath = @ScriptDir
		$sLocalFileName = 'test.zip'

		; 1 - do not show any confirmations dialogs;
		; 2 - add to queue only, do not start downloading.
		$sFlags = 2

		_IDM_SendLink($sURL, $sLocalPath, $sLocalFileName, $sFlags)
	EndIf
EndFunc   ;==>example
