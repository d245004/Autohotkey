#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1


InputBox,num,,�ݺ� ��ġ�� �Է��ϼ���. (MAX 999),,,,,,,,10
if (num <= 0 ){
	MsgBox,�Է� ��ġ�� Ȯ���ϼ���.
	Reload
}
if ErrorLevel
	ExitApp

Send,{ctrl down}{f2}{ctrl up}

^F2::
aa = 1
Loop, 1000
{	
	WorkWindow = WinGetActiveTitle,"[302]�繰���� - Chrome"
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
MsgBox, 4, �Ͻ� ����, �۾��� �Ϸ� �Ǿ����ϴ�.`n��� �۾��� �Ϸ��� "��" `n�۾��� �������� "�ƴϿ�" `nŰ�� ��������

IfMsgBox,yes
	Reload
IfMsgBox,no
	ExitApp

return

^Space::Pause   

^PGUP::Reload

^PGDN::ExitApp
