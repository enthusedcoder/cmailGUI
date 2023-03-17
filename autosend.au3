#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Add_Constants=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; *** Start added by AutoIt3Wrapper ***
#include <ListBoxConstants.au3>
; *** End added by AutoIt3Wrapper ***
; *** Start added by AutoIt3Wrapper ***
#include <AutoItConstants.au3>
; *** End added by AutoIt3Wrapper ***
; *** Start added by AutoIt3Wrapper ***
#include <MsgBoxConstants.au3>
; *** End added by AutoIt3Wrapper ***
#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.15.0 (Beta)
	Author:        William Higgs

	Script Function:
	Provides a graphical user interface for the command line utility "Cmail", which lets one send emails via command line.  I wrote this
	specificially to reduced the ammount of time needed to send messages to potential employers.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <Zip.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <ComboConstants.au3>
#include <InetConstants.au3>
#include <ListViewConstants.au3>
#include <File.au3>
#include <Array.au3>
#include <GuiEdit.au3>
#include <GuiListBox.au3>
#include <Constants.au3>
#include "GUIListViewEx.au3"
#include <Word.au3>
#include <Array.au3>
OnAutoItExitRegister("_Exit")
Global $trans = False
;_WordErrorHandlerRegister()
If Not FileExists ( @ScriptDir & "\CMail.exe" ) Then
	$fir = InetGet ( "https://www.inveigle.net/CMail_0.7.9b.zip", @ScriptDir & "\Cmail.zip", $INET_IGNORESSL )
	InetClose ( $fir )
	_Zip_UnzipAll ( @ScriptDir & "\Cmail.zip", @ScriptDir, 16 )
EndIf

If Not FileExists(@MyDocumentsDir & "\settings.ini") Then
	If Not IsDeclared("sInputBoxAnswer") Then Local $sInputBoxAnswer
	$sInputBoxAnswer = InputBox("Name", "What is the name wou want the recipients of your messages to see in regards to the sender? (Should be your name)", "", " ")
	If @error = 1 Then
		Exit
	Else
		IniWrite(@MyDocumentsDir & "\settings.ini", "Config", "Name", $sInputBoxAnswer)
	EndIf
	If Not IsDeclared("sInputBoxAnswer") Then Local $sInputBoxAnswer
	$sInputBoxAnswer = InputBox("Email address", "What is your email address or the email address of the account sending the mail?", "", " ")
	If @error = 1 Then
		Exit
	Else
		IniWrite(@MyDocumentsDir & "\settings.ini", "Config", "Email Address", $sInputBoxAnswer)
	EndIf
	If Not IsDeclared("sInputBoxAnswer") Then Local $sInputBoxAnswer
	$sInputBoxAnswer = InputBox("SMTP server", "What is the address of the SMTP server for your outgoing mail?", "", " ")
	If @error = 1 Then
		Exit
	Else
		IniWrite(@MyDocumentsDir & "\settings.ini", "Config", "SMTP Server", $sInputBoxAnswer)
	EndIf
	$sInputBoxAnswer = InputBox("Email Port", "What is the port used by smtp server?", "", " ")
	If @error = 1 Then
		Exit
	Else
		IniWrite(@MyDocumentsDir & "\settings.ini", "Config", "Port", $sInputBoxAnswer)
	EndIf
	$sInputBoxAnswer = InputBox("Email Username", "What is the username used to login to your email?", "", " ")
	If @error = 1 Then
		Exit
	Else
		IniWrite(@MyDocumentsDir & "\settings.ini", "Config", "User name", $sInputBoxAnswer)
	EndIf
	$sInputBoxAnswer = InputBox("Email Password", "What is the password used to login to your email?", "", " ")
	If @error = 1 Then
		Exit
	Else
		IniWrite(@MyDocumentsDir & "\settings.ini", "Config", "Password", $sInputBoxAnswer)
	EndIf
	If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
	$iMsgBoxAnswer = MsgBox($MB_YESNO + $MB_ICONQUESTION + $MB_SYSTEMMODAL, "Proxy server?", "Does your network utilize a proxy server?")
	If $iMsgBoxAnswer = $IDYES Then
		$sInputBoxAnswer = InputBox("Proxy server", "What is the proxy server address?", "", " ")
		If @error = 1 Then
			Exit
		Else
			IniWrite(@MyDocumentsDir & "\settings.ini", "Config", "Proxy", $sInputBoxAnswer)
		EndIf
	Else
		IniWrite(@MyDocumentsDir & "\settings.ini", "Config", "Proxy", "False")
	EndIf
EndIf
Global $oWordApp = _Word_Create(False)
Global $oDoc = _Word_DocAdd($oWordApp)
Global $oRange = $oDoc.Range
Global $oWordApp2 = _Word_Create(False)
Global $oDoc2 = _Word_DocAdd($oWordApp2)
Global $oRange2 = $oDoc2.Range
Global $oSpellCollection, $oAlternateWords, $oSpellCollection2, $oAlternateWords2
OnAutoItExitRegister("_Exit")
HotKeySet("^d", "stylish")
Global $attach = ""
Global $array
Global $both = False
Global $both2 = False
;$sphand = _Spell_HunspellInit ( "C:\hunspell-en_US\en_US.aff", "C:\hunspell-en_US\en_US.dic" )
$Form2 = GUICreate("Form2", 406, 514, 320, 200, -1, BitOR($WS_EX_ACCEPTFILES, $WS_EX_WINDOWEDGE))
$handle = WinGetHandle($Form2)
$MenuItem1 = GUICtrlCreateMenu("configure")
$MenuItem2 = GUICtrlCreateMenuItem("Attachments", $MenuItem1)
$Label1 = GUICtrlCreateLabel("Email Address", 128, 0, 130, 29)
GUICtrlSetFont(-1, 16, 400, 0, "MS Sans Serif")
$Input1 = GUICtrlCreateInput("", 70, 32, 273, 21)
$Label2 = GUICtrlCreateLabel("Attachments (Can drag and drop below)", 27, 64, 351, 29)
GUICtrlSetFont(-1, 16, 400, 0, "MS Sans Serif")
$Input2 = GUICtrlCreateInput("", 32, 104, 337, 21, $ES_READONLY)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
$Button9 = GUICtrlCreateButton("Add/Browse", 158, 136, 89, 33, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$Label4 = GUICtrlCreateLabel("Subject", 160, 184, 157, 29)
GUICtrlSetFont(-1, 16, 400, 0, "MS Sans Serif")
$Input3 = GUICtrlCreateInput("", 32, 216, 337, 21)
$Label3 = GUICtrlCreateLabel("Message", 160, 240, 172, 29)
GUICtrlSetFont(-1, 16, 400, 0, "MS Sans Serif")
$Edit1 = GUICtrlCreateEdit("", 32, 272, 345, 153, BitOR($ES_WANTRETURN, $WS_VSCROLL))
$Button1 = GUICtrlCreateButton("Send", 30, 440, 113, 33, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$Button2 = GUICtrlCreateButton("Save As Template", 160, 440, 105, 33, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$Button6 = GUICtrlCreateButton("Use Template", 282, 440, 113, 33, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$Form5 = GUICreate("Dynamic Variables", 274, 290, 750, 514)
$ListView1 = GUICtrlCreateListView("Dynamic Variable|Value", 15, 5, 241, 241, BitOR($LVS_SINGLESEL, $LVS_SHOWSELALWAYS) )
_GUICtrlListView_SetExtendedListViewStyle($ListView1, $LVS_EX_FULLROWSELECT)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 110)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 127)
$iLVIndex_1 = _GUIListViewEx_Init ( $ListView1 )
_GUIListViewEx_SetEditStatus($iLVIndex_1, 1, 1)

$Form1 = GUICreate("Spell Check", 396, 360, 694, 220)
Global $handle2 = WinGetHandle($Form1)
$Label7 = GUICtrlCreateLabel("Subject", 162, 8, 70, 29)
GUICtrlSetFont(-1, 15, 400, 0, "MS Sans Serif")
$Listbox3 = GUICtrlCreateList("", 24, 40, 137, 97, BitOR($LBS_NOTIFY, $LBS_EXTENDEDSEL, $LBS_HASSTRINGS, $WS_VSCROLL))
$Listbox4 = GUICtrlCreateList("", 232, 40, 137, 97, BitOR($LBS_NOTIFY, $LBS_EXTENDEDSEL, $LBS_HASSTRINGS, $WS_VSCROLL))
$Label8 = GUICtrlCreateLabel("Body", 173, 136, 49, 29)
GUICtrlSetFont(-1, 15, 400, 0, "MS Sans Serif")
$Listbox1 = GUICtrlCreateList("", 24, 168, 137, 97, BitOR($LBS_NOTIFY, $LBS_EXTENDEDSEL, $LBS_HASSTRINGS, $WS_VSCROLL))
$Listbox2 = GUICtrlCreateList("", 232, 168, 137, 97, BitOR($LBS_NOTIFY, $LBS_EXTENDEDSEL, $LBS_HASSTRINGS, $WS_VSCROLL))
$Label9 = GUICtrlCreateLabel("Errors", 64, 272, 53, 24)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
$Label10 = GUICtrlCreateLabel("Corrections", 256, 272, 96, 24)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
$Button3 = GUICtrlCreateButton("Send", 32, 313, 75, 25, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$Button5 = GUICtrlCreateButton("Correct Spelling", 152, 313, 83, 25, $BS_NOTIFY)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetCursor(-1, 0)
$Button4 = GUICtrlCreateButton("Cancel", 280, 313, 75, 25, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$Form3 = GUICreate("Form3", 286, 283, 192, 124)
$Label5 = GUICtrlCreateLabel("Attachments", 86, 8, 113, 29)
GUICtrlSetFont(-1, 15, 400, 0, "MS Sans Serif")
$List1 = GUICtrlCreateList("", 16, 40, 249, 188, $LBS_NOTIFY + $WS_VSCROLL)
$Button7 = GUICtrlCreateButton("Delete", 40, 240, 81, 33, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$Button8 = GUICtrlCreateButton("Back", 160, 240, 81, 33, $BS_NOTIFY)
GUICtrlSetCursor(-1, 0)
$Form4 = GUICreate("Form4", 242, 95, 216, 123)
$Combo1 = GUICtrlCreateCombo("", 40, 56, 169, 25, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
$Label6 = GUICtrlCreateLabel("Templates", 72, 8, 96, 29)
GUICtrlSetFont(-1, 15, 400, 0, "MS Sans Serif")
_GUIListViewEx_MsgRegister()
GUISetState(@SW_DISABLE, $Form4)
GUISetState(@SW_DISABLE, $Form3)
GUISetState(@SW_DISABLE, $Form1)
GUISetState(@SW_SHOW, $Form2)
$aRect = _GUICtrlEdit_GetRECT($Edit1)
$aRect[0] += 10
$aRect[1] += 10
$aRect[2] -= 10
$aRect[3] -= 10
_GUICtrlEdit_SetRECT($Edit1, $aRect)

While 1
	Global $nMsg = GUIGetMsg(1)
	Switch $nMsg[1]
		Case $Form2
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE
					_Exit()
				Case $GUI_EVENT_DROPPED
					Sleep(200)
					$mult = StringSplit(GUICtrlRead($Input2), "|")
					If @error Then
						WinActivate($Form2)
						WinWaitActive($Form2)
						GUICtrlSetState($Input2, $GUI_FOCUS)
						Send("{ENTER}", 0)
					Else
						For $v = 1 To $mult[0] Step 1
							$attach = $attach & $mult[$v] & ";"
						Next
						GUICtrlSetData($Input2, "")
						$sToolTipAnswer = ToolTip("The attachments were added!", Default, Default, "Success")
						Sleep(3000)
						ToolTip("")
					EndIf

				Case $MenuItem2
					GUISetState(@SW_HIDE, $Form2)
					GUISetState(@SW_DISABLE, $Form2)
					GUISwitch($Form3)
					GUISetState(@SW_ENABLE, $Form3)
					GUISetState(@SW_SHOW, $Form3)
					If $attach <> "" Then
						$array = StringSplit(StringTrimRight($attach, 1), ";")
						If @error Then
							SetError(0)
							_GUICtrlListBox_AddString($List1, _GetFilename(StringTrimRight($attach, 1)) & "." & _GetFilenameExt(StringTrimRight($attach, 1)))
						Else
							For $l = 1 To $array[0] Step 1
								_GUICtrlListBox_AddString($List1, _GetFilename($array[$l]) & "." & _GetFilenameExt($array[$l]))
							Next
						EndIf
					Else
						_GUICtrlListBox_AddString($List1, "You currently do not have any attachments.")
					EndIf
				Case $Input2
					$attach = $attach & GUICtrlRead($Input2) & ";"
					GUICtrlSetData($Input2, "")
					$sToolTipAnswer = ToolTip("The attachment was added!", Default, Default, "Success")
					Sleep(3000)
					ToolTip("")

				Case $Button9
					If GUICtrlRead($Input2) <> "" Then
						If FileExists(GUICtrlRead($Input2)) Then
							$attach = $attach & GUICtrlRead($Input2) & ";"
							GUICtrlSetData($Input2, "")
							$sToolTipAnswer = ToolTip("The attachment was added!", Default, Default, "Success")
							Sleep(3000)
							ToolTip("")
						Else
							If Not IsDeclared("sToolTipAnswer") Then Local $sToolTipAnswer
							$sToolTipAnswer = ToolTip("File name entered does not exist", Default, Default, "Failure", 0, 0)
							Sleep(3000)
							ToolTip("")
						EndIf

					Else
						$file = FileOpenDialog("Choose the file you want to attach.", "", "All (*.*)", 7, "", $Form2)
						If $file = "" Then

						Else
							$files = StringSplit($file, "|")
							If @error Then
								SetError(0)
								$attach = $attach & $file & ";"
								GUICtrlSetData($Input2, $file)
								$sToolTipAnswer = ToolTip("The attachment was added!", Default, Default, "Success")
								Sleep(1000)
								ToolTip("")
								GUICtrlSetData($Input2, "")
							Else
								For $p = 2 To $files[0] Step 1
									$attach = $attach & $files[1] & "\" & $files[$p] & ";"
								Next
								$sToolTipAnswer = ToolTip("The attachment was added!", Default, Default, "Success")
								Sleep(1000)
								ToolTip("")
							EndIf
						EndIf

					EndIf
				Case $Button1
					If StringRight(GUICtrlRead($Input3), 1) <> "." Or StringRight(GUICtrlRead($Input3), 1) <> "!" Or StringRight(GUICtrlRead($Input3), 1) <> "?" Or StringRight(GUICtrlRead($Input3), 1) <> " " Then
						GUICtrlSetData($Input3, GUICtrlRead($Input3) & " ")
					EndIf
					;_GUICtrlEdit_SetSel ( $Edit1, 0, -1 )
					_SpellCheck()
				Case $Button2 ;save as template
					If Not IsDeclared("sInputBoxAnswer") Then Local $sInputBoxAnswer
					$sInputBoxAnswer = InputBox("Template Name", "What do you want to name this template?", "", " ")
					If @error = 1 Then

					Else
						$name = $sInputBoxAnswer
						$num = Int(IniRead(@MyDocumentsDir & "\settings.ini", "Number", "Template", "0"))
						$num += 1
						$bodcap = _GUICtrlEdit_GetText($Edit1)
						$bodcap = StringReplace($bodcap, @CRLF, "\n")
						$subcap = GUICtrlRead($Input3)
						$attcap = $attach
						Local $holdarr[5][2] = [[3, ""], ["Name", $name], ["Body", $bodcap], ["Subject", $subcap], ["Attachments", $attcap]]
						IniWriteSection(@MyDocumentsDir & "\settings.ini", "Template " & $num, $holdarr)
						IniWrite(@MyDocumentsDir & "\settings.ini", "Number", "Template", $num)
						$sToolTipAnswer = ToolTip("The template was saved!", Default, Default, "Success")
						Sleep(2000)
						ToolTip("")
					EndIf

				Case $Button6
					GUISetState(@SW_HIDE, $Form2)
					GUISetState(@SW_DISABLE, $Form2)
					GUISwitch($Form3)
					GUISetState(@SW_ENABLE, $Form4)
					GUISetState(@SW_SHOW, $Form4)
					$numag = Int(IniRead(@MyDocumentsDir & "\settings.ini", "Number", "Template", "0"))
					If $numag = 0 Then
						MsgBox($MB_OK + $MB_ICONHAND, "No templates", "You have not configured any templates.")
					Else
						For $h = 1 To $numag Step 1
							GUICtrlSetData($Combo1, IniRead(@MyDocumentsDir & "\settings.ini", "Template " & $h, "Name", "NA"))
						Next
					EndIf

			EndSwitch
		Case $Form1
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_HIDE, $Form1)
					GUISetState(@SW_DISABLE, $Form1)
					GUISwitch($Form2)
					GUISetState(@SW_ENABLE, $Form2)
					GUISetState(@SW_SHOW, $Form2)
				Case $Listbox1
					;_SpellCheck2 ()
					_SpellingSuggestions()
				Case $Listbox2
					GUICtrlSetState($Button5, $GUI_ENABLE)
				Case $Listbox3
					;_SpellCheck ()
					_SpellingSuggestions2()
				Case $Listbox4
					GUICtrlSetState($Button5, $GUI_ENABLE)
				Case $Button3
					SendMessage()
				Case $Button4
					GUISetState(@SW_HIDE, $Form1)
					GUISetState(@SW_DISABLE, $Form1)
					GUISwitch($Form2)
					GUISetState(@SW_ENABLE, $Form2)
					GUISetState(@SW_SHOW, $Form2)
				Case $Button5
					If _GUICtrlListBox_GetSelCount($Listbox4) > 0 Then
						_ReplaceWord2()
					EndIf
					If _GUICtrlListBox_GetSelCount($Listbox2) > 0 Then
						_ReplaceWord()
					EndIf
					_SpellCheck()
			EndSwitch
		Case $Form3
			ToolTip("")
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE, $Button8
					ToolTip("")
					GUISetState(@SW_HIDE, $Form3)
					GUISetState(@SW_DISABLE, $Form3)
					GUISwitch($Form2)
					GUISetState(@SW_ENABLE, $Form2)
					GUISetState(@SW_SHOW, $Form2)
					If UBound($array) = 1 Then
						$attach = ""
					Else
						$attach = ""
						For $ff = 1 To UBound($array) - 1 Step 1
							$attach = $attach & $array[$ff] & ";"
						Next
					EndIf
					_GUICtrlListBox_ResetContent($List1)
				Case $List1
					ToolTip("")
				Case $Button7
					If _GUICtrlListBox_GetSelCount = -1 Then
						$sToolTipAnswer = ToolTip("Select something to delete first numbnuts.", Default, Default, "Idiot")
					Else
						ToolTip("")
						$seltext = _GUICtrlListBox_GetText($List1, _GUICtrlListBox_GetCurSel($List1))
						For $ff = 1 To $array[0] Step 1
							If StringInStr($array[$ff], $seltext) > 0 Then
								_ArrayDelete($array, $ff)
								_GUICtrlListBox_DeleteString($List1, _GUICtrlListBox_GetCurSel($List1))
								ExitLoop
							EndIf
						Next
					EndIf
			EndSwitch
		Case $Form4
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_HIDE, $Form4)
					GUISetState(@SW_DISABLE, $Form4)
					GUISwitch($Form2)
					GUISetState(@SW_ENABLE, $Form2)
					GUISetState(@SW_SHOW, $Form2)
				Case $Combo1
					$sec = ""
					$numag2 = Int(IniRead(@MyDocumentsDir & "\settings.ini", "Number", "Template", "0"))
					If $numag2 = 0 Then

					Else
						$use = GUICtrlRead($Combo1)
						For $h = 1 To $numag2 Step 1
							If StringCompare(IniRead(@MyDocumentsDir & "\settings.ini", "Template " & $h, "Name", "NA"), $use) = 0 Then
								$sec = "Template " & $h
								ExitLoop
							Else
								ContinueLoop
							EndIf
						Next
						GUISetState(@SW_HIDE, $Form4)
						GUISetState(@SW_DISABLE, $Form4)
						GUISwitch($Form2)
						GUISetState(@SW_ENABLE, $Form2)
						GUISetState(@SW_SHOW, $Form2)
						$attach = IniRead(@MyDocumentsDir & "\settings.ini", $sec, "Attachments", "NA")
						GUICtrlSetData($Input3, IniRead(@MyDocumentsDir & "\settings.ini", $sec, "Subject", "NA"))
						$thetext = IniRead(@MyDocumentsDir & "\settings.ini", $sec, "Body", "NA")
						_GUICtrlEdit_SetText($Edit1, StringReplace($thetext, "\n", @CRLF))
						$subcheck = True
						$bodcheck = True
						$loc1 = ""
						$loc2 = ""
						$hol = 1
						While 1
							$varcheck = StringInStr ( IniRead(@MyDocumentsDir & "\settings.ini", $sec, "Subject", "NA"), "%", Default, $hol )
							If $varcheck = 0 Or @error Then
								SetError ( 0 )
								ExitLoop
							Else
								$hol += 1
								If StringMid ( IniRead(@MyDocumentsDir & "\settings.ini", $sec, "Subject", "NA"), Int ( $varcheck ) - 1, 1 ) <> "\" Then
									$loc1 = $loc1 & $varcheck & ";"
								Else
									ContinueLoop
								EndIf
							EndIf
						WEnd
						If $loc1 <> "" Then
							$subcheck = True
							$holsplit = StringSplit ( StringTrimRight ( $loc1, 1 ), ";", $STR_NOCOUNT )
							$values = ""
							For $i = 0 To UBound ( $holsplit ) - 1 Step 2
								$values = $values & StringMid ( IniRead(@MyDocumentsDir & "\settings.ini", $sec, "Subject", "NA"), Int ( $holsplit[$i] ), ( Int ( $holsplit[$i+1] ) - Int ( $holsplit[$i] ) ) + 1 ) & ";"
							Next
						Else
							$subcheck = False
						EndIf


						$hol = 1
						While 1
							$varcheck2 = StringInStr ( StringReplace($thetext, "\n", @CRLF), "%", Default, $hol )
							If $varcheck2 = 0 Or @error Then
								SetError ( 0 )
								ExitLoop
							Else
								$hol += 1
								If StringMid ( StringReplace($thetext, "\n", @CRLF), Int ( $varcheck2 ) - 1, 1 ) <> "\" Then
									$loc2 = $loc2 & $varcheck2 & ";"
								Else
									ContinueLoop
								EndIf
							EndIf
						WEnd
						If $loc2 <> "" Then
							$bodcheck = True
							$holsplit = StringSplit ( StringTrimRight ( $loc2, 1 ), ";", $STR_NOCOUNT )
							$values2 = ""
							For $i = 0 To UBound ( $holsplit ) - 1 Step 2
								$values2 = $values2 & StringMid ( StringReplace($thetext, "\n", @CRLF), Int ( $holsplit[$i] ), ( Int ( $holsplit[$i+1] ) - Int ( $holsplit[$i] ) ) + 1 ) & ";"
							Next
						Else
							$bodcheck = False
						EndIf



						If $subcheck Or $bodcheck Then
							GUISetState ( @SW_SHOW, $Form5 )
							_GUIListViewEx_SetActive ( $iLVIndex_1 )
							If $subcheck Then
								$holsplit = StringSplit ( StringTrimRight ( $values, 1 ), ";", $STR_NOCOUNT )
								For $i = 0 To UBound ( $holsplit ) - 1 Step 1
									Global $arrr[2] = [StringTrimLeft ( StringTrimRight ( $holsplit[$i], 1 ), 1 ), $holsplit[$i]]
									_GUIListViewEx_Insert ( $arrr )
									;GUICtrlCreateListViewItem ( $holsplit[$i] & "|" & StringTrimLeft ( StringTrimRight ( $holsplit[$i], 1 ), 1 ), $ListView1 )
								Next
							EndIf
							If $bodcheck Then
								$holsplit = StringSplit ( StringTrimRight ( $values2, 1 ), ";", $STR_NOCOUNT )
								For $i = 0 To UBound ( $holsplit ) - 1 Step 1
									Global $arrr[2] = [StringTrimLeft ( StringTrimRight ( $holsplit[$i], 1 ), 1 ), $holsplit[$i]]
									_GUIListViewEx_Insert ( $arrr )
									;GUICtrlCreateListViewItem ( $holsplit[$i] & "|" & StringTrimLeft ( StringTrimRight ( $holsplit[$i], 1 ), 1 ), $ListView1 )
								Next
							EndIf
							$ind = _GUIListViewEx_ReadToArray ( $ListView1 )
							$iLVIndex_1 = _GUIListViewEx_Init ( $ListView1, $ind )
						_GUIListViewEx_SetEditStatus($iLVIndex_1, 1, 1)
						;GUISwitch ( $Form5 )
						EndIf
					EndIf
			EndSwitch
		Case $Form5
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE
					GUISetState ( @SW_HIDE, $Form5 )
			EndSwitch

	EndSwitch
	$vRet = _GUIListViewEx_EventMonitor()
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Error", "Event error: " & @error)
	EndIf
	Switch @extended
		Case 1
			If $vRet <> "" Then
				$row = $vRet[1][0]
				$col = $vRet[1][1]
				$before = $vRet[1][2]
				$after = $vRet[1][3]
				$rep = StringReplace ( GUICtrlRead ( $Input3 ), $before, $after, Default, $STR_CASESENSE )
				GUICtrlSetData ( $Input3, $rep )
				$rep2 = StringReplace ( _GUICtrlEdit_GetText ( $Edit1 ), $before, $after, Default, $STR_CASESENSE )
				_GUICtrlEdit_SetText ( $Edit1, $rep2 )
			EndIf
	EndSwitch
WEnd
Func stylish()
	$thestyle = GUIGetStyle($handle)
	If $trans = False Then
		GUISetStyle(-1, $thestyle[1] + 32, $handle)
		WinSetTrans($handle, "", 170)
		$trans = True
	Else
		GUISetStyle(-1, $thestyle[1] - 32, $handle)
		WinSetTrans($handle, "", 255)
		$trans = False
	EndIf
EndFunc   ;==>stylish
Func GetHoveredHwnd()
	Local $iRet = DllCall("user32.dll", "int", "WindowFromPoint", "long", MouseGetPos(0), "long", MouseGetPos(1))
	If IsArray($iRet) Then Return HWnd($iRet[0])
	Return SetError(1, 0, 0)
EndFunc   ;==>GetHoveredHwnd
Func _SpellCheck()
	Local $sText, $sText2, $sWord, $sWord2

	$both = False
	$pText = _GUICtrlEdit_GetText($Edit1)
	$oRange.Delete
	$oRange.InsertAfter($pText)
	_SetLanguage()
	$both2 = False
	$pText2 = GUICtrlRead($Input3)
	$oRange2.Delete
	$oRange2.InsertAfter($pText2)
	_SetLanguage2()

	$oSpellCollection = $oRange.SpellingErrors
	$oSpellCollection2 = $oRange2.SpellingErrors
	;Local $argg[$oSpellCollection.Count]
	_GUICtrlListBox_ResetContent($Listbox1)
	_GUICtrlListBox_ResetContent($Listbox2)
	_GUICtrlListBox_ResetContent($Listbox3)
	_GUICtrlListBox_ResetContent($Listbox4)
	GUICtrlSetState($Button5, $GUI_DISABLE)
	If $oSpellCollection.Count > 0 Or $oSpellCollection2.Count > 0 Then
		If BitAND(WinGetState($handle), 2) Then
			If $trans = True Then
				stylish()
			EndIf
			GUISetState(@SW_DISABLE, $Form2)
			GUISetState(@SW_HIDE, $Form2)
			GUISwitch($Form1)
			GUISetState(@SW_SHOW, $Form1)
			GUISetState(@SW_ENABLE, $Form1)
		EndIf

		If $oSpellCollection.Count > 0 Then

			For $i = 1 To $oSpellCollection.Count
				_GUICtrlListBox_AddString($Listbox1, $oSpellCollection.Item($i).Text)
			Next
			;_GUICtrlEdit_SetText($Edit1, $oRange.Text)
			;GUICtrlSetData($Input3, $oRange2.Text)
		Else
			_GUICtrlListBox_AddString($Listbox1, "No spelling errors")
		EndIf
		If $oSpellCollection2.Count > 0 Then
			_GUICtrlListBox_ResetContent($Listbox3)
			_GUICtrlListBox_ResetContent($Listbox4)
			GUICtrlSetState($Button5, $GUI_DISABLE)
			For $p = 1 To $oSpellCollection2.Count
				_GUICtrlListBox_AddString($Listbox3, $oSpellCollection2.Item($p).Text)
			Next
			;_GUICtrlEdit_SetText ($Edit1, $oRange.Text)
			;GUICtrlSetData($Input3, $oRange2.Text)
		Else
			_GUICtrlListBox_AddString($Listbox3, "No spelling errors")
		EndIf

	Else
		_GUICtrlEdit_SetText ($Edit1, $oRange.Text)
		GUICtrlSetData($Input3, $oRange2.Text)
		SendMessage()
	EndIf

EndFunc   ;==>_SpellCheck


Func _SpellingSuggestions()
	Local $iWord, $ssWord
	;
	_GUICtrlListBox_ResetContent($Listbox2)
	If _GUICtrlListBox_GetSelCount($Listbox4) = 0 Then
		GUICtrlSetState($Button5, $GUI_DISABLE)
	EndIf



	$iWord = _GUICtrlListBox_GetCurSel($Listbox1) + 1
	$ssWord = $oSpellCollection.Item($iWord).Text
	$oAlternateWords = $oWordApp.GetSpellingSuggestions($ssWord)



	If $oAlternateWords.Count > 0 Then
		For $v = 1 To $oAlternateWords.Count
			_GUICtrlListBox_AddString($Listbox2, $oAlternateWords.Item($v).Name)
		Next
	Else
		_GUICtrlListBox_AddString($Listbox2, "No suggestions.")
	EndIf
EndFunc   ;==>_SpellingSuggestions

Func _SpellingSuggestions2()
	Local $iWord2, $ssWord2
	;
	_GUICtrlListBox_ResetContent($Listbox4)
	If _GUICtrlListBox_GetSelCount($Listbox2) = 0 Then
		GUICtrlSetState($Button5, $GUI_DISABLE)
	EndIf


	$iWord2 = _GUICtrlListBox_GetCurSel($Listbox3) + 1
	$ssWord2 = $oSpellCollection2.Item($iWord2).Text
	$oAlternateWords2 = $oWordApp2.GetSpellingSuggestions($ssWord2)



	If $oAlternateWords2.Count > 0 Then
		For $c = 1 To $oAlternateWords2.Count
			_GUICtrlListBox_AddString($Listbox4, $oAlternateWords2.Item($c).Name)
		Next
	Else
		_GUICtrlListBox_AddString($Listbox4, "No suggestions.")
	EndIf
EndFunc   ;==>_SpellingSuggestions2

Func _ReplaceWord()
	Local $iWord, $iNewWord, $sWord, $sNewWord, $sText, $sNewText
	;
	$iWord = _GUICtrlListBox_GetCurSel($Listbox1) + 1
	$iNewWord = _GUICtrlListBox_GetCurSel($Listbox2) + 1
	If $iWord == $LB_ERR Or $iNewWord == $LB_ERR Then
		;MsgBox(48, "Error", "You must first select a word to replace, then a replacement word.")
		;Return
	Else
		$oSpellCollection.Item($iWord).Text = $oAlternateWords.Item($iNewWord).Name
		_GUICtrlEdit_SetText($Edit1, $oRange.Text)
	EndIf


EndFunc   ;==>_ReplaceWord
Func _ReplaceWord2()
	Local $iWord2, $iNewWord2, $sWord2, $sNewWord2, $sText2, $sNewText2
	;
	$iWord2 = _GUICtrlListBox_GetCurSel($Listbox3) + 1
	$iNewWord2 = _GUICtrlListBox_GetCurSel($Listbox4) + 1
	If $iWord2 == $LB_ERR Or $iNewWord2 == $LB_ERR Then
		;MsgBox(48, "Error", "You must first select a word to replace, then a replacement word.")
		;Return
	Else
		$oSpellCollection2.Item($iWord2).Text = $oAlternateWords2.Item($iNewWord2).Name
		GUICtrlSetData($Input3, $oRange2.Text)
	EndIf


EndFunc   ;==>_ReplaceWord2

Func _SetLanguage()
	$sLang = "English"
	$oWordApp.CheckLanguage = False
	$WdLangID = Number(1033)

	If $WdLangID Then
		With $oRange
			.LanguageID = $WdLangID
			.NoProofing = False
		EndWith

	EndIf
EndFunc   ;==>_SetLanguage
Func _SetLanguage2()
	$sLang = "English"
	$oWordApp2.CheckLanguage = False
	$WdLangID = Number(1033)
	If $WdLangID Then
		With $oRange2
			.LanguageID = $WdLangID
			.NoProofing = False
		EndWith
	EndIf
EndFunc   ;==>_SetLanguage2

Func SendMessage()
	If BitAND(WinGetState($handle2), 2) Then
		GUISetState(@SW_HIDE, $Form1)
		GUISetState(@SW_DISABLE, $Form1)
		GUISwitch($Form2)
		GUISetState(@SW_ENABLE, $Form2)
		GUISetState(@SW_SHOW, $Form2)
	EndIf
	$change = StringStripWS(GUICtrlRead($Input1), 3)
	GUICtrlSetData($Input1, $change)
	$body = _GUICtrlEdit_GetText ($Edit1)
	;_GUICtrlEdit_SetText($Edit1, "")
	;$attach = StringStripCR ( $attach )
	$subject = StringStripWS ( GUICtrlRead($Input3), 3 )
	If $attach <> "" Then
		$split = StringSplit(StringTrimRight($attach, 1), ';')
		If @error Then
			SetError(0)
			$finattach = '-a:"' & StringTrimRight ( $attach, 1 ) & '" '
		Else
			$finattach = Null
			For $i = 1 To $split[0] Step 1
				$finattach = $finattach & '-a:"' & $split[$i] & '" '
			Next
		EndIf
	EndIf
	$body = StringReplace($body, @CRLF, "\n")
	$body = StringReplace ( $body, '"', '\"' )
	#cs
		$split = StringSplit ( $body, "", $STR_NOCOUNT )
		For $i = 0 To UBound ( $split ) - 1 Step 1
		If $split[$i] = Chr ( 10 ) Then
		$split[$i] = "\n"
		ElseIf $split[$i] = Chr ( 13 ) Then
		$split[$i] = Null
		Else
		ContinueLoop
		EndIf
		Next
		$body = _ArrayToString ( $split, "" )
		ConsoleWrite ( $body )
	#ce
	If StringRegExp(GUICtrlRead($Input1), "^[A-Za-z0-9](([_\.\-]?[a-zA-Z0-9]+)*)@([A-Za-z0-9]+)(([\.\-]?[a-zA-Z0-9]+)*)\.([A-Za-z]{2,})$") = 1 Then
		If $attach = "" Then
			$proc = Run('CMail.exe -from:' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Email Address", "NA") & ':"' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Name", "NA") & '" -to:' & GUICtrlRead($Input1) & ' -subject:"' & $subject & '" -body:"' & $body & '" -host:' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "User name", "NA") & ':' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Password", "NA") & '@' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "SMTP Server", "NA") & ':' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Port", "NA") & ' -starttls -requiretls -d', @ScriptDir, @SW_SHOW, $STDOUT_CHILD)
		Else
			$proc = Run('CMail.exe -from:' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Email Address", "NA") & ':"' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Name", "NA") & '" -to:' & GUICtrlRead($Input1) & ' -subject:"' & $subject & '" -body:"' & $body & '" ' & $finattach & '-host:' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "User name", "NA") & ':' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Password", "NA") & '@' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "SMTP Server", "NA") & ':' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Port", "NA") & ' -starttls -requiretls -d', @ScriptDir, @SW_SHOW, $STDOUT_CHILD)
			ClipPut('CMail -from:' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Email Address", "NA") & ':"' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Name", "NA") & '" -to:' & GUICtrlRead($Input1) & ' -subject:"' & $subject & '" -body:"' & $body & '" ' & $finattach & '-host:' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "User name", "NA") & ':' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Password", "NA") & '@' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "SMTP Server", "NA") & ':' & IniRead(@MyDocumentsDir & "\settings.ini", "Config", "Port", "NA") & ' -starttls -requiretls -d')
			ProcessWaitClose($proc)
			$text = StdoutRead($proc)
			ClipPut($text)
		EndIf

			MsgBox($MB_OK + $MB_SYSTEMMODAL, "Sent!", "Message Sent!!")
	Else
		MsgBox(16, "Not a valid email", "Enter a correct email address please")
	EndIf
EndFunc   ;==>SendMessage

Func _GetFilename($sFilePath)
	Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
	Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
	If IsObj($oColFiles) Then
		For $oObjectFile In $oColFiles
			Return $oObjectFile.FileName
		Next
	EndIf
	Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilename

Func _GetFilenameExt($sFilePath)
	Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
	Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
	If IsObj($oColFiles) Then
		For $oObjectFile In $oColFiles
			Return $oObjectFile.Extension
		Next
	EndIf
	Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilenameExt

Func _GetFilenameInt($sFilePath)
	Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
	Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
	If IsObj($oColFiles) Then
		For $oObjectFile In $oColFiles
			Return $oObjectFile.Name
		Next
	EndIf
	Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilenameInt

Func _GetFilenameDrive($sFilePath)
	Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
	Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
	If IsObj($oColFiles) Then
		For $oObjectFile In $oColFiles
			Return StringUpper($oObjectFile.Drive)
		Next
	EndIf
	Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilenameDrive

Func _GetFilenamePath($sFilePath)
	Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
	Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
	If IsObj($oColFiles) Then
		For $oObjectFile In $oColFiles
			Return $oObjectFile.Path
		Next
	EndIf
	Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilenamePath

Func _Exit()
	_Word_Quit($oWordApp)
	_Word_Quit($oWordApp2)
	Exit
EndFunc   ;==>_Exit
