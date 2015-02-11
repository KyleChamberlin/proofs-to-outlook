#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Version=Beta
#AutoIt3Wrapper_Icon=ProofsEmail.ico
#AutoIt3Wrapper_Outfile=proofs.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=0.2.0.2
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <File.au3>
#include <ComboConstants.au3>
#include <string.au3>
#include <Array.au3>
#include <GuiConstantsEx.au3>
#include <AVIConstants.au3>
#include <TreeViewConstants.au3>
#include <Debug.au3>

Global $settingsfile = @TempDir & "\proofSettings.txt"
Global $NameLine = 1
Global $TitleLine = 2
Global $ExtLine = 3
FileInstall("./emailsig.png", @TempDir & "\emailsig.png", 1)

Global $file = ""

If $cmdline[0] = "" Then
	$file = FileOpenDialog("Please Select a Proof File", "c:\spool\out\", "PDFs (*.pdf)|All (*.*)")
	If $file = "" Then
		Exit
	EndIf
Else
	$file = $cmdline[1]
EndIf

Global $settings = FileOpen($settingsfile, 0)
$name = FileReadLine($settings, $NameLine)
$title = FileReadLine($settings, $TitleLine)
$ext = FileReadLine($settings, $ExtLine)
FileClose($settings)
If $name = "" Then $name = "Your Name Here"
If $title = "" Then $title = "Your Title Here"
If $ext = "" Then $ext = "00"

Dim $path_drive, $path_folder, $path_file, $path_ext, $AE, $AM, $client_email, $client_name, $Job_Name
_PathSplit($file, $path_drive, $path_folder, $path_file, $path_ext)

$return = FileCopy($file, @TempDir & "\" & $path_file & $path_ext, 1)

$file = @TempDir & "\" & $path_file & $path_ext

$bodyText1 = 'I have attached proofs of '
$bodyText2 = '. Please review your proofs for accuracy and correctness. Once you are satisfied with the quality of your proofs, please reply with approval to proceed with the mailing.'
$bodytext2use = $bodyText2
$size = FileGetSize($file) / 1048576

If $size > 9.75 Then
	$attach = 0
	$url = sendToShareFile($file, $path_file & $path_ext)
	$bodyText1 = 'I have uploaded proofs of '
	$linkAnchor = '<a href="' & $url & '">click here</a>'
	$bodytext2use = 'To view your proofs please click here, and download the file. Then simply review your proofs for accuracy and correctness. Once you are satisfied with the quality of your proofs, please reply with approval to proceed with the mailing.'
	$bodyText2 = '. To view your proofs please ' & $linkAnchor & ', and download the file. Then simply review your proofs for accuracy and correctness. Once you are satisfied with the quality of your proofs, please reply with approval to proceed with the mailing.'
Else
	$attach = 1
EndIf

$gUID = GUICreate("Proofs to Email")
$client_emailGUI = GUICtrlCreateInput("send.proofs@here.com", 64, 6, 121, 21)
GUICtrlCreateLabel("To:", 8, 8, 36, 17)
GUICtrlCreateLabel("CC:", 8, 48, 36, 17)
$AMGUI = GUICtrlCreateCombo("Account Manager", 64, 46, 145, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
GUICtrlSetData($AMGUI, "Ericka|Lisa|Autumn")
$AEGUI = GUICtrlCreateCombo("Account Executive", 272, 46, 145, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
GUICtrlSetData($AEGUI, "Natascha|Matt|Brian")
GUICtrlCreateLabel("CC:", 224, 48, 36, 17)
$subjectGUI = GUICtrlCreateInput($path_file, 64, 80, 300, 21)
GUICtrlCreateLabel("Subject:", 8, 80, 36, 17)
$client_nameGUI = GUICtrlCreateInput("Client Name", 8, 108, 121, 21)
GUICtrlCreateLabel($bodyText1, 8, 152)
$JnameGUI = GUICtrlCreateInput($path_file, 8, 176, 300, 21)
GUICtrlCreateLabel($bodytext2use, 8, 208, 460, 65)
$nameGUI = GUICtrlCreateInput($name, 8, 280, 222, 21)
GUICtrlCreateLabel("/", 240, 282, 20, 65)
$titleGUI = GUICtrlCreateInput($title, 261, 280, 200, 21)
GUICtrlCreateLabel("616.957.2120 ext. ", 8, 320, 100, 17)
$extGUI = GUICtrlCreateInput($ext, 109, 318, 121, 21)
$done = GUICtrlCreateButton("Submit", 8, 360, 75, 25)
GUISetState(@SW_SHOW)

$loop = 1
While $loop = 1
	$nMsg = GUIGetMsg()
	If $nMsg = $done Then
		$loop = 0
		$client_name = GUICtrlRead($client_nameGUI)
		$client_email = GUICtrlRead($client_emailGUI)
		$AM = GUICtrlRead($AMGUI)
		$AE = GUICtrlRead($AEGUI)
		$Jobname = GUICtrlRead($JnameGUI)
		$name = GUICtrlRead($nameGUI)
		$title = GUICtrlRead($titleGUI)
		$ext = GUICtrlRead($extGUI)
		$sSubj = GUICtrlRead($subjectGUI)
	EndIf
	If $nMsg = $GUI_EVENT_CLOSE Then
		Exit
	EndIf
WEnd

_FileWriteToLine($settingsfile, $NameLine, $name, 1)
_FileWriteToLine($settingsfile, $TitleLine, $title, 1)
_FileWriteToLine($settingsfile, $ExtLine, $ext, 1)

GUIDelete($gUID)

Switch $AM
	Case "Ericka"
		$AM = "erickac@kentcommunications.com"
	Case "Lisa"
		$AM = "lisan@kentcommunications.com"
	Case "Autumn"
		$AM = "autumnh@kentcommunications.com"
	Case $AE
		$AE = ""
	Case "Account Manager"
		$AM = ""
EndSwitch

Switch $AE
	Case "Matt"
		$AE = "mattb@kentcommunications.com"
	Case "Natascha"
		$AE = "natascham@kentcommunications.com"
    Case $AE = "Brian"
		$AE = "brianq@kentcommunications.com"
	Case "Account Executive"
		$AE = ""
EndSwitch

Func sendToShareFile($fileToSend, $filename)
    ; TODO: implement Sharefile API and return URL
    ; temp prompt for url from sharefile
    Return InputBox("Sharefile URL","Enter the sharefile URL below","")
EndFunc   ;==>sendToShareFile

$sRcpt = $client_email
$sCCRcpt = $AM & "; " & $AE
$sBody = $client_name & ',<p><p>' & $bodyText1 & $Jobname & $bodyText2 & '<p><p>Thanks,<p><p><b><span style=' & Chr(39) & 'color:#4F81BD' & Chr(39) & '>' & $name & '</span></b><span style=' & Chr(39) & 'color:#A6A6A6' & Chr(39) & '> /' & $title & '<br></span><span style=' & Chr(39) & 'color:#A6A6A6' & Chr(39) & '>616.957.2120 ext. ' & $ext & '<br></span><span style=' & Chr(39) & 'color:#474847' & Chr(39) & '><img width=147 height=76 id="Picture_x0020_1" src="' & @TempDir & '\emailsig.png" alt="cid:image001.gif@01CBE224.635165D0"></span><span style=' & Chr(39) & 'color:#A6A6A6' & Chr(39) & '></span>'

$outlook = ObjCreate("outlook.application")
If $outlook = 0 Then
	$outlook = ObjGet("outlook.application")
EndIf
$mail = ObjCreate('outlook.mailitem')
$mail = $outlook.CreateItem(0)

$mail.To = $sRcpt
$mail.cc = $sCCRcpt
If $attach = 1 Then
	$mail.attachments.add($file)
EndIf
$mail.Subject = $sSubj
$mail.BodyFormat = 2
$mail.HTMLBody = $sBody
$mail.Display
