#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True

^F2::
InputBox,num,���� �� ����,���� ���� �Է� �Ͻÿ�. `n�⺻ ������ 2�� �Է��Ͻÿ�.


Loop 1000
{
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1
	while (tmp1)
	{
		VAR2=E%num%
		VAR6=F%num%
		part := Xl.Range(VAR2).value

		if (part = "" )
		{
			; MsgBox,�۾��� ���� �մϴ�
			; ExitApp
			MsgBox, ��� �۾� �Ϸ��� F2 Ű�� �����ÿ�
			Return
		}
		Click,384,113,2
		Sleep,500
		send %part%
		send {enter}
		sleep 1500     ;   1000�� 1�ʸ� �ǹ��Ѵ�

		Clipboard=
            MouseClickDrag,L,88,234,168,240 ;ENGLISH
            Sleep,500
            Send ^c
            clipwait,0
            Xl.range(var6).value := Clipboard



		Sleep 1500
		num++
		;Pause   ;��� ����
	}
}


return    ; ���α׷� ����ġ��


^Space::Pause   ; ������ ��� ������ ON / OFF ��Ȱ�� �Ѵ�

^PGDN::ExitApp   ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�

^PgUp::Reload