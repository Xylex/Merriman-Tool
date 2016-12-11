;general usful functions tool written by Vincent Merriman

#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <ScreenCapture.au3>

#cs

#ce

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;START FUNCTIONS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;internet conection function
Func _IsInternetConnected()
    Local $aReturn = DllCall('connect.dll', 'long', 'IsInternetConnected')
    If @error Then
        Return SetError(1, 0, False)
    EndIf
    Return $aReturn[0] = 0
EndFunc   ;==>_IsInternetConnected
;end internet connection function

;check internet explorer version
Func _GetIEVersion()
    Return StringRegExpReplace(RegRead('HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\', 'Version'), '^(\d+\.\d+).*', '$1')
EndFunc   ;==>_GetIEVersion
;end internet explorer version check

;get time zone
Func _GetTimeOnline($iTimeZone)
	Local $aTimeZone[7] = ['utc', 'est', 'cst', 'mst', 'pst', 'akst', 'hast']

	Local $sRead = BinaryToString(InetRead('http://www.timeapi.org/' & $aTimeZone[$iTimeZone] & '/now?format=\Y/\m/\d%20\H:\M:\S'))

	If @error Then
		Return SetError(1, 0, @YEAR & '/' & @MON & '/' & @MDAY & ' ' & @HOUR & ':' & @MIN & ':' & @SEC)
	EndIf

	Return $sRead
EndFunc   ;==>_GetTimeOnline
;end get time zone

;get ssid start
Func _GetActiveSSID()
    Local $iPID = Run(@ComSpec & ' /u /c ' & 'netsh wlan show interfaces', @SystemDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD), $sOutput = ''
    While 1
        $sOutput &= StdoutRead($iPID)
        If @error Then
            ExitLoop
        EndIf
        $sOutput = StringStripWS($sOutput, 7)
    WEnd

    $sReturn = StringRegExp($sOutput, '(?s)(?i)SSID\s*:\s(.*?)' & @CR, 3)
    If @error Then
        Return SetError(1, 0, '')
    EndIf
    Return $sReturn[0]
EndFunc   ;==>_GetActiveSSID
;getssid end

;is admin enabled
Func _IsAdminEnabled()
    Local $oWMIService = ObjGet('winmgmts:\\localhost\root\CIMV2')
    Local $oColItems = $oWMIService.ExecQuery('SELECT * FROM Win32_UserAccount WHERE Name = "Administrator"', "WQL", 0x30)
    If IsObj($oColItems) Then
        For $oItem In $oColItems
            Return $oItem.Disabled = False
        Next
    EndIf
    Return True
EndFunc   ;==>_IsAdminEnabled
;end is admin enabled

;windows update check and automation
Func _WindowsUpdate()
    Return Run(@ComSpec & ' /c wuauclt /a /DetectNow', @SystemDir, @SW_HIDE)
EndFunc   ;==>_WindowsUpdate
;windows update check and automation

; Cancel printer jobs for the default printer or the printer name provided.

Func _CancelPrinterJobs($sPrinterName = '')
    If StringStripWS($sPrinterName, 8) = '' Then
        $sPrinterName = 'Default = True'
    Else
        $sPrinterName = 'Name = "' & $sPrinterName & '"'
    EndIf
    Local $iResult = 0, $oWMIService = ObjGet('winmgmts:\\' & '.' & '\root\cimv2')
    Local $oColItems = $oWMIService.ExecQuery('Select * From Win32_Printer Where ' & $sPrinterName)
    If IsObj($oColItems) Then
        For $oObjectItem In $oColItems
            $iResult = $oObjectItem.CancelAllJobs()
        Next
    EndIf
    Return $iResult
EndFunc   ;==>_CancelPrinterJobs
;end cancel print jobs

;os info dump
Func _GetOSVersion()
    Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Local $colSettings = $objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For $objOperatingSystem In $colSettings
        Return "Windows " & StringMid($objOperatingSystem.Caption, 19)
    Next
EndFunc   ;==>_getOSVersion
;end info dump

;is the system a vm?
Func _IsVirtualMachine() ; Returns True or False.
    Local $oColItems, $oWMIService
    $oWMIService = ObjGet("winmgmts:\\localhost\root\cimv2")
    $oColItems = $oWMIService.ExecQuery("Select * From Win32_ComputerSystemProduct", "WQL", 0x30)
    If IsObj($oColItems) Then
        For $oObjectItem In $oColItems
            Return StringRegExp($oObjectItem.Name, 'VirtualBox|VMWare|Virtual PC') = 1
        Next
    EndIf
    Return False
EndFunc   ;==>_IsVirtualMachine
;end vm check

;drive check
Func DriveCheck()
	Local $aDriveArray = DriveGetDrive("ALL")
	If @error = 0 Then
		Local $sDriveInfo = ""
		For $i = 1 To $aDriveArray[0]
			$sDriveInfo &= StringUpper($aDriveArray[$i]) & "\" & @CRLF
			$sDriveInfo &= @TAB & "File System = " & DriveGetFileSystem($aDriveArray[$i]) & @CRLF
			$sDriveInfo &= @TAB & "Label = " & DriveGetLabel($aDriveArray[$i]) & @CRLF
			$sDriveInfo &= @TAB & "Serial = " & DriveGetSerial($aDriveArray[$i]) & @CRLF
			$sDriveInfo &= @TAB & "Type = " & DriveGetType($aDriveArray[$i]) & @CRLF
			$sDriveInfo &= @TAB & "Free Space = " & DriveSpaceFree($aDriveArray[$i]) & @CRLF
			$sDriveInfo &= @TAB & "Total Space = " & DriveSpaceTotal($aDriveArray[$i]) & @CRLF
			$sDriveInfo &= @TAB & "Status = " & DriveStatus($aDriveArray[$i]) & @CRLF
			$sDriveInfo &= @CRLF
		Next
		MsgBox(4096, "", $sDriveInfo)
	EndIf
EndFunc   ;==>Example
;end drive check

;usb device check
Func _UsbCheck()
	$strComputer = "."
	$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\cimv2")
	$colDevices = $objWMIService.ExecQuery ("Select * From Win32_USBControllerDevice")
	For $objDevice in $colDevices
		$strDeviceName = $objDevice.Dependent
		;ConsoleWrite("!>" & $strDeviceName & @CRLF)
		$strQuotes = Chr(34)
		$strDeviceName = StringReplace($strDeviceName, $strQuotes, "")
		$arrDeviceNames = StringSplit($strDeviceName, "=")
		$strDeviceName = $arrDeviceNames[2]
		$colUSBDevices = $objWMIService.ExecQuery ("Select * From Win32_PnPEntity Where DeviceID = '" & $strDeviceName & "'")
		For $objUSBDevice in $colUSBDevices
			MsgBox(0, "USB Device Info Dump", "-->" & $objUSBDevice.Description & @CRLF)
		Next
	Next
EndFunc
;end usb device check

;get computer model information (returns null if the pc is custom built)
Global $aArray = _ComputerNameAndModel() ; Returns an Array with 2 indexes.
;MsgBox(64, "_ComputerNameAndModel()", 'The Product is a "' & $aArray[0] & '" and the Serial Number is "' & $aArray[1] & '".')

Func _ComputerNameAndModel()
    Local $aReturn[2] = ["(Unknown)", "(Unknown)"], $oColItems, $oWMIService

    $oWMIService = ObjGet("winmgmts:\\.\root\cimv2")
    $oColItems = $oWMIService.ExecQuery("Select * From Win32_ComputerSystemProduct", "WQL", 0x30)
    If IsObj($oColItems) Then
        For $oObjectItem In $oColItems
            $aReturn[0] = $oObjectItem.Name
            $aReturn[1] = $oObjectItem.IdentifyingNumber
        Next
        Return $aReturn
    EndIf
    Return SetError(1, 0, $aReturn)
EndFunc   ;==>_ComputerNameAndModel
;end get computer model information

;kill all active windows
Func _KillActiveWindows()
    $Window = WinGetTitle("[ACTIVE]")
    $CheckCancel = MsgBox(4, "Kill...", "Do you wish to kill the following process?" & @CRLF & @CRLF & $Window)
    If $CheckCancel = 6 Then
        $Process = WinGetProcess($Window)
        ProcessClose($Process)
    Else
        Sleep(200)
    EndIf
EndFunc   ;==>_Test
;end kill all active windows

;open edge
Func _DoEdge()
	Local $edge_test = FileExists(@WindowsDir & '\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe')
                    If $edge_test = 1 Then
                        Local $edge = RunWait('explorer.exe shell:AppsFolder\Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge')
                    EndIf
EndFunc
;end open microsoft edge

;needed mem var
Local $aMem = MemGetStats()
;end

;start print screen function
Func _PrintScrn()
    ; Capture full screen
    _ScreenCapture_Capture(@MyDocumentsDir & "\GDIPlus_Image1.jpg")
    ShellExecute(@MyDocumentsDir & "\GDIPlus_Image1.jpg")
EndFunc
;end print screen function

;get active CD rom drives
Func _GetCD()
$DISK = DriveGetDrive("ALL")

For $I = 1 To $DISK[0]

    $CDROM = DriveGetType($DISK[$I])
    If $CDROM = "CDROM" Then
        CDTray($DISK[$I], "closed")
        Sleep(500) ; wait for CD to close
    EndIf


    $STATUS = DriveStatus($DISK[$I])
    If $STATUS = "ready" Then
        MsgBox(0,"Disc Drive Status",">DRIVE = " & $DISK[$I] & "      STATUS = " & $STATUS & @CRLF)

    Else
        MsgBox(0,"Disc Drive Status","DRIVE = " & $DISK[$I] & "      STATUS = " & $STATUS & @CRLF)
    EndIf
Next
EndFunc
;end get cd rom drives

;start image resize functions
Func _ResizeImage()
	Local $sImageName = InputBox("Original Image", "Enter the name of the image to resize, plus extention (example is image.jpg).")
	Local $sImageWidth = InputBox("Image Width", "Enter the width for the new image.")
	Local $sImageHeight = InputBox("Image Height", "Enter the height for the new image.")
	_ImageResize(@ScriptDir & "\" & $sImageName, @ScriptDir & "\" & "RESIZED" & $sImageName, $sImageWidth, $sImageHeight)
EndFunc

Func _ImageResize($sInImage, $sOutImage, $iW, $iH)
    Local $hWnd, $hDC, $hBMP, $hImage1, $hImage2, $hGraphic, $CLSID, $i = 0

    ;OutFile path, to use later on.
    Local $sOP = StringLeft($sOutImage, StringInStr($sOutImage, "\", 0, -1))

    ;OutFile name, to use later on.
    Local $sOF = StringMid($sOutImage, StringInStr($sOutImage, "\", 0, -1) + 1)

    ;OutFile extension , to use for the encoder later on.
    Local $Ext = StringUpper(StringMid($sOutImage, StringInStr($sOutImage, ".", 0, -1) + 1))

    ; Win api to create blank bitmap at the width and height to put your resized image on.
    $hWnd = _WinAPI_GetDesktopWindow()
    $hDC = _WinAPI_GetDC($hWnd)
    $hBMP = _WinAPI_CreateCompatibleBitmap($hDC, $iW, $iH)
    _WinAPI_ReleaseDC($hWnd, $hDC)

    ;Start GDIPlus
    _GDIPlus_Startup()

    ;Get the handle of blank bitmap you created above as an image
    $hImage1 = _GDIPlus_BitmapCreateFromHBITMAP ($hBMP)

    ;Load the image you want to resize.
    $hImage2 = _GDIPlus_ImageLoadFromFile($sInImage)

    ;Get the graphic context of the blank bitmap
    $hGraphic = _GDIPlus_ImageGetGraphicsContext ($hImage1)

    ;Draw the loaded image onto the blank bitmap at the size you want
    _GDIPLus_GraphicsDrawImageRect($hGraphic, $hImage2, 0, 0, $iW, $iW)

    ;Get the encoder of to save the resized image in the format you want.
    $CLSID = _GDIPlus_EncodersGetCLSID($Ext)

    ;Generate a number for out file that doesn't already exist, so you don't overwrite an existing image.
    Do
        $i += 1
    Until (Not FileExists($sOP & $i & "_" & $sOF))

    ;Prefix the number to the begining of the output filename
    $sOutImage = $sOP & $i & "_" & $sOF

    ;Save the new resized image.
    _GDIPlus_ImageSaveToFileEx($hImage1, $sOutImage, $CLSID)

    ;Clean up and shutdown GDIPlus.
    _GDIPlus_ImageDispose($hImage1)
    _GDIPlus_ImageDispose($hImage2)
    _GDIPlus_GraphicsDispose ($hGraphic)
    _WinAPI_DeleteObject($hBMP)
    _GDIPlus_Shutdown()
EndFunc
;end image resize functions

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;END FUNCTIONS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;define output strings (ip addy and shit) CRLF is new line in autoit and & is equivilent to + to combine strings
local $sIpData = "IP of first network adapter: " & @IPAddress1 & @CRLF & "IP of the second network adapter (if applicable): " & @IPAddress2 & @CRLF & "IP of the third network adapter (if applicable): " & @IPAddress3 & @CRLF & "IP of the fourth network adapter (if applicable): " & @IPAddress4
local $sDriveData = "Home directory drive letter is: " & @HomeDrive & @CRLF & "Directory part it: " & @HomePath & @CRLF & "Server and share name containing current user's home directory (if applicable): " & @HomeShare
local $sAppDataPath = "App_Data file path is: " & @LocalAppDataDir
local $sOsInfo = "OS Type: " & @OSType & @CRLF & "OS Version: " & @OSVersion & @CRLF & "OS Architecture: " & @OSArch & @CRLF & "OS Build: " & @OSBuild & @CRLF & "OS Language: " & @OSLang & @CRLF & "Service Pack: " & @OSServicePack
local $sDNSinfo = "Log On DNS Domain: " & @LogonDNSDomain & @CRLF & "Log on Domain: " & @LogonDomain & @CRLF & "Log on server: " & @LogonServer
Local $iRecycle

;GUI SHITTERY
GUICreate("Merri Tool", 250, 500)
GUISetState()
;$CTRL_btn1 = GUICtrlCreateButton("1", 54, 138, 36, 29)
Local $CTRL_btn1 = GUICtrlCreateButton("IP Info", -1, -1, 36, 29)
Local $CTRL_btn2 = GUICtrlCreateButton("Drive Info", 36, -1, 50, 29)
Local $CTRL_btn3 = GUICtrlCreateButton("App Data Path", 86, -1, 85, 29)
Local $CTRL_btn4 = GUICtrlCreateButton("OS", 171, -1, 25, 29)
Local $CTRL_btn5 = GUICtrlCreateButton("DNS Info", 196, -1, 50, 29)
Local $CTRL_btn6 = GUICtrlCreateButton("BEEP", 0, 29, 35, 29);sound test
Local $CTRL_btn7 = GUICtrlCreateButton("Empty Trash", 35, 29, 85, 29);
Local $CTRL_btn8 = GUICtrlCreateButton("Documents", 120, 29, 65, 29);
Local $CTRL_btn9 = GUICtrlCreateButton("Programs", 185, 29, 65, 29);
Local $CTRL_btn10 = GUICtrlCreateButton("Local App Data", 0, 58, 85, 29);
Local $CTRL_btn11 = GUICtrlCreateButton("Roaming App Data", 85, 58, 95, 29);
Local $CTRL_btn12 = GUICtrlCreateButton("Reset Browser", 0, 87, 80, 29);
Local $CTRL_btn13 = GUICtrlCreateButton("Minimize All", 80, 87, 65, 29);
Local $CTRL_btn14 = GUICtrlCreateButton("Notepad", 180, 58, 50, 29);
Local $CTRL_btn15 = GUICtrlCreateButton("Internet?", 145, 87, 50, 29);
Local $CTRL_btn16 = GUICtrlCreateButton("IE Version#", 0, 116, 65, 29);
Local $CTRL_btn17 = GUICtrlCreateButton("Time", 65, 116, 35, 29);
Local $CTRL_btn18 = GUICtrlCreateButton("SSID", 100, 116, 35, 29);
Local $CTRL_btn19 = GUICtrlCreateButton("CMD", 135, 116, 30, 29);
Local $CTRL_btn20 = GUICtrlCreateButton("Admin Status", 165, 116, 70, 29);
Local $CTRL_btn21 = GUICtrlCreateButton("Windows Update", 0, 145, 90, 29);
Local $CTRL_btn22 = GUICtrlCreateButton("Stop Print Jobs", 90, 145, 80, 29);
Local $CTRL_btn23 = GUICtrlCreateButton("VM?", 170, 145, 30, 29);
Local $CTRL_btn24 = GUICtrlCreateButton("USB?", 200, 145, 35, 29);
Local $CTRL_btn25 = GUICtrlCreateButton("Computer Model?", 0, 174, 90, 29);
Local $CTRL_btn26 = GUICtrlCreateButton("Calculator", 90, 174, 55, 29);
Local $CTRL_btn27 = GUICtrlCreateButton("Kill Active Processes", 145, 174, 105, 29);
Local $CTRL_btn28 = GUICtrlCreateButton("Open Edge", 0, 203, 60, 29);
Local $CTRL_btn29 = GUICtrlCreateButton("RAM?", 60, 203, 35, 29);
Local $CTRL_btn30 = GUICtrlCreateButton("Run I.E.", 95, 203, 45, 29);
Local $CTRL_btn31 = GUICtrlCreateButton("Run Firefox", 140, 203, 60, 29);
Local $CTRL_btn32 = GUICtrlCreateButton("Print Screen", 0, 232, 65, 29);
Local $CTRL_btn33 = GUICtrlCreateButton("Close Disc Drive", 65, 232, 85, 29);
Local $CTRL_btn34 = GUICtrlCreateButton("Resize Image", 150, 232, 70, 29);

GUISetState()
;END GUI SHITTERY

;GET BUTTON PRESS INPUT
  Local $msg
Do
    $msg = GUIGetMsg()
	;MsgBox(0,"debug",$msg)
	if $msg == $CTRL_btn1 Then MsgBox(0, "IP information", $sIpData)
	if $msg == $CTRL_btn2 Then MsgBox(0, "Drive Info", DriveCheck())
	if $msg == $CTRL_btn3 Then MsgBox(0, "App_Data Path", $sAppDataPath)
	if $msg == $CTRL_btn4 Then MsgBox(0, "Operating System Information", $sOsInfo & @CRLF & _GetOSVersion())
	if $msg == $CTRL_btn5 Then MsgBox(0, "DNS Info", $sDNSinfo)
	if $msg == $CTRL_btn6 Then Beep(500, 1000)
	if $msg == $CTRL_btn7 Then $iRecycle = FileRecycleEmpty(@HomeDrive)& MsgBox(0,"Recycle Bin Status","Recycle Bin Emptied...");
	if $msg == $CTRL_btn8 Then Run("explorer.exe " & @MyDocumentsDir)
	if $msg == $CTRL_btn9 Then Run("explorer.exe " & @ProgramsDir)
	if $msg == $CTRL_btn10 Then Run("explorer.exe " & @LocalAppDataDir)
	if $msg == $CTRL_btn11 Then Run("explorer.exe " & @AppDataDir)
	if $msg == $CTRL_btn12 Then RunWait ( 'rundll32.exe inetcpl.cpl ResetIEtoDefaults', '', @SW_HIDE )
	if $msg == $CTRL_btn13 Then WinMinimizeAll()
	if $msg == $CTRL_btn14 Then Run("notepad")
	if $msg == $CTRL_btn15 Then MsgBox(0, "Internet Status", "Internet Is Connected" & " = " & _IsInternetConnected() & @CRLF)
	if $msg == $CTRL_btn16 Then MsgBox(0, "I.E. Version Check", "I.E. Version: " & _GetIEVersion() & @CRLF)
	if $msg == $CTRL_btn17 Then MsgBox(0, "Online Time Stamp", "Time from online servers (UTC Time): " & _GetTimeOnline(0) & @CRLF)
	if $msg == $CTRL_btn18 Then MsgBox(0, "SSID Information", "SSID Info Dump: " & _GetActiveSSID() & @CRLF)
	if $msg == $CTRL_btn19 Then Run("cmd.exe")
	if $msg == $CTRL_btn20 Then MsgBox(0, "Admin Status", _IsAdminEnabled() & @CRLF)
	if $msg == $CTRL_btn21 Then MsgBox(0, "Win Updates", "Update Check: " & _WindowsUpdate() & @CRLF)
	if $msg == $CTRL_btn22 Then MsgBox(0, "Print Status", "Print job request(s) stopped." & _CancelPrinterJobs() & @CRLF)
	if $msg == $CTRL_btn23 Then MsgBox(0, "Is Running OS A VM?", "VM Status: " & _IsVirtualMachine() & @CRLF)
	if $msg == $CTRL_btn24 Then _UsbCheck()
	if $msg == $CTRL_btn25 Then MsgBox(64, "_ComputerNameAndModel()", 'The Product is a "' & $aArray[0] & '" and the Serial Number is "' & $aArray[1] & '".')
	if $msg == $CTRL_btn26 Then Run("calc.exe")
	if $msg == $CTRL_btn27 Then _KillActiveWindows()
	if $msg == $CTRL_btn28 Then _DoEdge()
	if $msg == $CTRL_btn29 Then MsgBox($MB_SYSTEMMODAL, "", "Total physical RAM (GB): " & ($aMem[1] * 0.001) * 0.001)
	if $msg == $CTRL_btn30 Then Run(@ProgramFilesDir & "\Internet Explorer\IEXPLORE.EXE")
	if $msg == $CTRL_btn31 Then Run(@ProgramFilesDir & "\Mozilla Firefox\firefox.exe")
	if $msg == $CTRL_btn32 Then _PrintScrn()
	if $msg == $CTRL_btn33 Then _GetCD()
	if $msg == $CTRL_btn34 Then _ResizeImage()

Until $msg = $GUI_EVENT_CLOSE
;GUICreate("Merri Tool", 250, 500)