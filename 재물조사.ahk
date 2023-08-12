#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1

^F2::
InputBox,num,작업 수량 입력,작업하고자 하는 수량을 입력합시다
num := num+1
Loop
{
	aa = 1
	Loop
	{
		ImageSearch, vx, vy, 740,388,1013,466,*250 find_ok.png
		vx := vx + 100
		vy := vy + 45
		if (ErrorLevel=0)
			Click,%vx%,%vy%
			;~ MsgBox, find
		if (ErrorLevel=1)
			;~ Continue
			MsgBox, 염병 못찾는다. 확인 바람
		Sleep,500
		Send,{right}{enter}
		Sleep,500
		Click,%vx%,%vy%
		Sleep,5000
		;~ aa++
		aa += 1
		If (num <= aa)
		{
			MsgBox, 작업완료! 계속 하려면 Ctrl+F2 누르시요
			Exit
		}
	}
	Pause
}
return

^Space::Pause

^PGDN::ExitApp

