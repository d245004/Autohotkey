;Coordmode, Mouse, Screen   ;  ��üȭ�鿡�� ���콺 ����Ʈ�� ����Ѵٰ� �����ϴ� ��
num=2
XLS_file_path3 := A_WorkingDir . "���� ��� �ͽ���Ʈ.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:G").NUMBERFORMAT := "@"
return

^F2::
Loop 1000
{
	WinGetActiveTitle,tmp1
		while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
	{
		VAR2=B%num%   ;    B���� �� ��°
		part := X1.Range(VAR2).value
		VAR3=E%num%
		QTY := round(X1.range(var3).value,0)
		var5=J%num%


		if part =             ; PART ������ null�̸�
			{
				MsgBox,�����մϴ�
				ExitApp
			}
			Click,1261,207
			Sleep 1300
			Click,209,336
			Sleep 1300
			Send %part%
			Sleep 1300
			Send {Tab}
			Sleep 1300
			Send %qty%
			Sleep 1300
			Click,1385,206
			Send 1300
			Send {Enter}
			Sleep 1000
			X1.Range(var5).value := "OK"

		num++

		Sleep 1700
		;~ Pause   ;��� ����
}
Pause  ;��� ����
}

return    ; ���α׷� ����ġ��

^PGUP::Reload

^Space::Pause

^PGDN::ExitApp   ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�