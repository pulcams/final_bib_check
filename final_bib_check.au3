#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_outfile=cat_processing.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;=================================================================
; NOTE: This is NOT for member copy sorting.
; A quick check of member copy in Voyager Cataloging module for:
; 1. 050 090 not P (and not PS 8001-8599)
; 2. 300 has a number other than 1
; 3. There's at least one 6xx_0 (if no 050 090 P)
; Using Voyager 9.x
; From 20150901
; pmg
;=================================================================

Local $f050a_count = 0
Local $flag = 0
Local $lcsh = False
Local $num300 = False
Local $cm300 = False
Local $msg = ""
Local $pcallno = False
Local $relator = True

WinActivate("Voyager Cataloging")
Send("!m")

;====================================
; enter record
;====================================
Send("{shiftdown}{tab 2}{shiftup}")
Send("{home}^{home}")
Send("{f8}")
Send("^c")
Sleep(100)
$thisfield = ClipGet()

While $thisfield <> "" And $thisfield < "711" ; loop through 7xx or to the end of the record
	;===============================
	; 050
	;===============================
	; Cannot be P or PS8001-8599
	If $thisfield == "050" Then
		Send("{TAB 3}{HOME}+{END}")
		Send("^c")
		Sleep(100)
		$f050 = ClipGet()
		If StringRegExp($f050, "‡a P(S\s+8[0-5][0-9][0-9])?") Then
			$pcallno = True
		EndIf
		Send("+{TAB 3}")
	EndIf
	;===============================
	; 090
	;===============================
	; Cannot be P or PS8001-8599
	If $thisfield == "090" Then
		Send("{TAB 3}{HOME}+{END}")
		Send("^c")
		Sleep(100)
		$f090 = ClipGet()
		If StringRegExp($f090, "‡a P(S\s+8[0-5][0-9][0-9])?") Then
			$pcallno = True
		EndIf
		Send("+{TAB 3}")
	EndIf
	;===============================
	; 300
	;===============================
	; Test for numbers other than 1
	If $thisfield == "300" Then
		Send("{TAB 3}{HOME}+{END}")
		Send("^c")
		Sleep(100)
		$f300 = ClipGet()
		If StringRegExp($f300, "[^1][0-9]+") Then
			$num300 = True
		EndIf
		If StringRegExp($f300,"\scm\.?") Then
			$cm300 = True
		EndIf
		Send("+{TAB 3}")
	EndIf
	;===============================
	; 6xx
	;===============================
	; check for at least one 6xx_0
	If StringMid($thisfield, 1, 1) == "6" Then
		Send("{TAB 2}")
		Send("^c")
		Sleep(100)
		$ind2 = ClipGet()
		If $ind2 == "0" Then
			$lcsh = True
		EndIf
		Send("+{TAB 2}")
	EndIf
	;===============================
	; 7xx
	;===============================
	; check that 700 or 710 has $e
	If  $thisfield == "700" or  $thisfield == "710" Then
		Send("{TAB 3}{HOME}+{END}")
		Send("^c")
		Sleep(100)
		$f7xx = ClipGet()
		If not StringRegExp($f7xx, ".*‡e\s.*") Then
			$relator = False
		EndIf
		Send("+{TAB 3}")
	EndIf

	ClipPut("") ; clear clipboard

	Send("{down}{f8}") ; keep moving

	Send("^c")
	Sleep(100) ; this may need to be tweaked

	$thisfield = ClipGet()
WEnd

;===============================
; Check variables
;===============================
If $pcallno == False and $lcsh == False Then
	$msg = "No P call no. and no 6XX_0"
	$flag = 48
EndIf

If $num300 == False Then
	$msg = $msg & @CRLF & "No number > 1 in 300 field"
	$flag = 48
EndIf

If $cm300 == False Then
	$msg = $msg & @CRLF & "No 'cm.' in 300 field"
	$flag = 48
EndIf

If $relator == False Then
	$msg = $msg & @CRLF & "No $e in 700 or 710"
	$flag = 48
EndIf

If $msg == "" Then
	$msg = "OK"
EndIf

;===============================
; Alert the hominid
;===============================
MsgBox($flag, "", $msg)