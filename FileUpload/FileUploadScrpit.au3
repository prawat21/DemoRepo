#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Dibyanshu Rai

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

$count = 0

Sleep(3000)


   $chrome = WinActivate("Open")


   if $chrome <> 0 Then
	  ControlFocus("Open","","Edit1")
	  Sleep(500)
	  ControlSetText("Open","","Edit1",$CmdLine[1])
	  Sleep(500)
	  ControlClick("Open","","Button1")
	  Exit



   EndIf


