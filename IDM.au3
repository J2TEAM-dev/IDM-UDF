#include-once

; #INDEX# =======================================================================================================================
; Title .........: Internet Download Manager (IDM)
; UDF Version....: 1.0.0
; AutoIt Version : 3.3.14.2
; Description ...: Download files using IDM COM based API
; Author(s) .....: Juno_okyo
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
Global $IDM_CLSID, $IDM_IID, $IDM_Desc, $IDM_INTERFACE = Default
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================
Global Const $IDM_DEBUG = False
Global Const $IDM_Init_Status = __IDM_Init()
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _IDM_SendLink
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __IDM_Init
; __IDM_ScanIID
; ===============================================================================================================================

If $IDM_DEBUG Then
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $IDM_CLSID = ' & $IDM_CLSID & @CRLF & '>Error code: ' & @error & @CRLF)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $IDM_IID = ' & $IDM_IID & @CRLF & '>Error code: ' & @error & @CRLF)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $IDM_Desc = ' & $IDM_Desc & @CRLF & '>Error code: ' & @error & @CRLF)
EndIf

; #FUNCTION# ====================================================================================================================
; Author ........: Juno_okyo
; ===============================================================================================================================
Func _IDM_SendLink($sURL, $sLocalPath = '', $sLocalFileName = '', $iFlags = 0, $sReferer = '', $sCookies = '', $sPostData = '', $sUsername = '', $sPassword = '')
	; Flags can be zero or a combination of the following values:
	; 1 - do not show any confirmations dialogs;
	; 2 - add to queue only, do not start downloading.

	If $IDM_INTERFACE = Default Then
		$IDM_INTERFACE = ObjCreateInterface($IDM_CLSID, $IDM_IID, $IDM_Desc)
	EndIf

	$IDM_INTERFACE.SendLinkToIDM($sURL, $sReferer, $sCookies, $sPostData, $sUsername, $sPassword, $sLocalPath, $sLocalFileName, $iFlags)
EndFunc   ;==>_IDM_SendLink

#Region <INTERNAL_USE_ONLY>
Func __IDM_Init($createInterface = False)
	Local $sKeyName = 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes'

	$IDM_CLSID = RegRead($sKeyName & '\IDMan.CIDMLinkTransmitter\CLSID', '')

	If Not @error And StringLen($IDM_CLSID) > 0 Then
		; Default value
		$IDM_IID = '{4BD46AAE-C51F-4BF7-8BC0-2E86E33D1873}'

		$sKeyName &= '\Interface\'

		; Check IID
		Local $sValue = RegRead($sKeyName & $IDM_IID, '')

		If @error Then
			$IDM_IID = __IDM_ScanIID($sKeyName)

			; Read again
			$sValue = RegRead($sKeyName & $IDM_IID, '')
		EndIf

		If $sValue = StringLower('ICIDMLinkTransmitter') Then
			$IDM_Desc = 'SendLinkToIDM HRESULT(BSTR; BSTR; BSTR; BSTR; BSTR; BSTR; BSTR; BSTR; long)'

			; Auto create Interface
			If $createInterface Then
				$IDM_INTERFACE = ObjCreateInterface($IDM_CLSID, $IDM_IID, $IDM_Desc)
			EndIf

			Return True
		EndIf
	EndIf

	Return SetError(1, 0, False)
EndFunc   ;==>__IDM_Init

Func __IDM_ScanIID($sKeyName)
	Local $sSubKey, $sValue, $i = 1

	Do
		$sSubKey = RegEnumKey($sKeyName, $i)
		If @error Then ExitLoop

		$sValue = RegRead($sKeyName & '\' & $sSubKey, '')

		If $IDM_DEBUG Then ConsoleWrite('> ' & $i & ' - ' & $sSubKey & ' - ' & $sValue & @CRLF)

		; Counter
		$i += 1
	Until $sValue = StringLower('ICIDMLinkTransmitter')

	Return $sSubKey
EndFunc   ;==>__IDM_ScanIID
#EndRegion <INTERNAL_USE_ONLY>
