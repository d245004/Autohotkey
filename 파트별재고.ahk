;Coordmode, Mouse, Screen   ;  ��üȭ�鿡�� ���콺 ����Ʈ�� ����Ѵٰ� �����ϴ� ��
num=2
XLS_file_path3 := A_WorkingDir . "HR07.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:E").NUMBERFORMAT := "@"
return

F2::
Loop 1000
{	
	WinGetActiveTitle,tmp1
		while (tmp1 = "��Ʈ����� - Whale")
	{
		VAR2=C%num%   ;    C���� �� ��°
		var3=E%num%
		part := X1.Range(VAR2).value
		MouseClick, left,108,146
		Send, {SHIFTDOWN}{END}{SHIFTUP}{DEL}
		send %part%
		send {enter}
		sleep 1000     ;   1000�� 1�ʸ� �ǹ��Ѵ�
		
		if
			MouseClick,L,215,400   ;  ǰ���� H �϶�
		else
			MouseClick,L,215,420   ;  ǰ���� K �϶�
		Sleep,1000
		MouseClick,L,208,361
		Sleep,1000
		MouseClick,L,271,691
		Sleep,1000
		WinGetActiveTitle, tmp2
			while (tmp2 = "������� - Whale")
			{
				qty := round(X1.range(var3).value,0)
		        send %qty%
				send {enter}
				Sleep,1000
				send {enter}
				Sleep,1000
				MouseClick,L,219,426
				Exit
			}
		num++
		Pause   ;��� ����
}
Pause  ;��� ����
}

return    ; ���α׷� ����ġ��

PGUP::Pause   ; ������ ��� ������ ON / OFF ��Ȱ�� �Ѵ�

PGDN::ExitApp   ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�