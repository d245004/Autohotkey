#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1

^F2::
InputBox,num,�۾� ���� �Է�,�۾��ϰ��� �ϴ� ������ �Է��սô�
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
			MsgBox, ���� ��ã�´�. Ȯ�� �ٶ�
		Sleep,500
		Send,{right}{enter}
		Sleep,500
		Click,%vx%,%vy%
		Sleep,5000
		;~ aa++
		aa += 1
		If (num <= aa)
		{
			MsgBox, �۾��Ϸ�! ��� �Ϸ��� Ctrl+F2 �����ÿ�
			Exit
		}
	}
	Pause
}
return

^Space::Pause

^PGDN::ExitApp

