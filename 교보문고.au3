#NoTrayIcon
#AutoIt3Wrapper_Change2CUI=y
#include <Array.au3>
#include <IE.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>
APPM_WEB('APPM_WEB','-','192.168.2.78','https://www.kyobobook.co.kr/login/login.laf','8443','*','*','-','yes','yes','-')
Func APPM_WEB($hostname,$sid,$ipaddr,$url,$application_port,$accountid,$password,$exe_path,$auto_logon_flag,$auto_submit_flag,$setforground_flag)
   ;MsgBox($MB_OKCANCEL, "Form Info", $ipaddr)   ;;https://www.autoitscript.com/autoit3/docs/functions/MsgBox.htm

   $url = $url

   ShellExecute(@ProgramFilesDir & "\internet explorer\iexplore.exe",$url)

   Local $count = 0

   while 1
	  $sTitle = WinGetTitle("[active]") ;;활성화되어있는 윈도우 창의 타이틀을 취득함
	  if $sTitle = "로그인 - 인터넷교보문고 - Internet Explorer" Then  ;;활성화되어있는 윈도우 창의 타이틀이 저와 같으면 반복문 끝냄
		 ExitLoop
	  EndIf
	  Sleep(10)
	  if $count > 6000 Then
		 MsgBox(1, "Form Info", $count)
		 Exit(-1)
	  Else
		 $count = $count + 1
	  EndIf
   WEnd

   $oIE = _IEAttach("", "embedded") ;브라우저에 연결
   _IELoadWait($oIE) ;브라우저가 로드될때까지 기다림

   While _IEPropertyGet($oIE, "busy") = True
	   Sleep(50)
   WEnd

   Local $oForm = _IEFormGetObjByName($oIE,"loginForm")
 MsgBox(0,"",$oForm)
   IF @error Then
	  Return
   EndIf

   Local $oUsername = _IEFormElementGetObjByName($oForm, "memid")
   Local $oPassword = _IEFormElementGetObjByName($oForm, "pw")

   _IEFormElementSetValue($oUsername, $accountid)
   _IEFormElementSetValue($oPassword, $password)
   send('{ENTER}')



      while 1
	  $sTitle = WinGetTitle("[active]") ;;활성화되어있는 윈도우 창의 타이틀을 취득함
	  if $sTitle = "꿈을 키우는 세상 - 인터넷교보문고 - Internet Explorer" Then  ;;활성화되어있는 윈도우 창의 타이틀이 저와 같으면 반복문 끝냄
		 ExitLoop
	  EndIf
	  Sleep(10)
	  if $count > 6000 Then
		 MsgBox(1, "Form Info", $count)
		 Exit(-1)
	  Else
		 $count = $count + 1
	  EndIf
   WEnd
      $oIE = _IEAttach("", "embedded") ;브라우저에 연결
   _IELoadWait($oIE) ;브라우저가 로드될때까지 기다림
   $tags = $oIE.document.GetElementsByTagName("a")
   For $tag in $tags
	   $class_value = $tag.GetAttribute("href")
	   ;MsgBox($MB_SYSTEMMODAL, "Form Info", $class_value)
	   If $class_value = "http://www.kyobobook.co.kr/event/dailyCheckSpci.laf?orderClick=c1j" Then
		 $tag.click()
		 ExitLoop
	   EndIf
	Next
      while 1
	  $sTitle = WinGetTitle("[active]") ;;활성화되어있는 윈도우 창의 타이틀을 취득함
	  if $sTitle = "출석체크 x 문장수집 - 인터넷교보문고 - Internet Explorer" Then  ;;활성화되어있는 윈도우 창의 타이틀이 저와 같으면 반복문 끝냄
		 ExitLoop
	  EndIf
	  Sleep(10)
	  if $count > 6000 Then
		 MsgBox(1, "Form Info", $count)
		 Exit(-1)
	  Else
		 $count = $count + 1
	  EndIf
   WEnd

   $oIE2 = _IEAttach("", "embedded") ;브라우저에 연결
   _IELoadWait($oIE2) ;브라우저가 로드될때까지 기다림
Local $ops = _IETagNameGetCollection($oIE2, "p")
For $op In $ops
   $a = $op.innertext
    If $op.classname == "part" Then
        MsgBox($MB_SYSTEMMODAL, "Form Info",$a)
        ExitLoop
    EndIf
 Next




 Local $oIE2 = _IECreate("www.google.com")
   _IELoadWait($oIE2) ;브라우저가 로드될때까지 기다림
      $oIE2 = _IEAttach("", "embedded") ;브라우저에 연결
 Local $oForm2 = _IEFormGetObjByName($oIE2,"f")
 Local $osearch = _IEFormElementGetObjByName($oForm2, "q")
  _IEFormElementSetValue($osearch, $a)
send('{ENTER}')




   $oqs = $oIE2.document.GetElementsByTagName("div"))
   
for $oq IN $oqs

	  if $oq.classname == "rc" Then
		 $b = $oq.innertext
		 if $b = StringSTR($b, "문장수집" Then
			$b.click
			ExitLoop
		 EndIf
	  EndIf
   Next




EndFunc


