#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1


InputBox,num,,반복 수치를 입력하세요. (MAX 999),,,,,,,,10
if (num <= 0 ){
	MsgBox,입력 수치를 확인하세요.
	Reload
}
if ErrorLevel
	ExitApp

Send,{ctrl down}{f2}{ctrl up}

^F2::
aa = 1
Loop, 1000
{	
	WorkWindow = WinGetActiveTitle,"[302]재물조사 - Chrome"
	while (WorkWindow)
	{
		Loop,1000 
		{
			Click,388,437
			Sleep,300
			Click,380,411
			Sleep,300
			Click,420,151
			Sleep,300
			Click,388,437
			Sleep,300
			Click,380,411
			Sleep,700
			aa++
			if (aa > num)
				break
		}
		break
	}
break
}
MsgBox, 4, 일시 중지, 작업이 완료 되었습니다.`n계속 작업을 하려면 "예" `n작업을 끝내려면 "아니요" `n키를 누르세요

IfMsgBox,yes
	Reload
IfMsgBox,no
	ExitApp

return

^Space::Pause   

^PGUP::Reload

^PGDN::ExitApp
