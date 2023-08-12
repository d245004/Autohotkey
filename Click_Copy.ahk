#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1




#^LButton::
Loop
{
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
    while (tmp1)
	{
		Click,,,2
		Sleep,300
		Send ^c
		sleep,300
		;MsgBox %Clipboard%
		sleep,300
		send ^2
		return	
	}
}


!LButton::
Loop
{
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
    while (tmp1)
	{
		Click
		sleep,300
		loop,50{
			send,{BackSpace}
			;sleep,10
		}
		
		send,^v
		Sleep,300
		send,{enter}
		Sleep,300
		return	
	}
}




^!PGDN::ExitApp

       