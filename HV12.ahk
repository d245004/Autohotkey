;Coordmode, Mouse, Screen   ;  ��üȭ�鿡�� ���콺 ����Ʈ�� ����Ѵٰ� �����ϴ� ��
num=110
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
		VAR3=I%num%
		QTY := round(X1.range(var3).value,0)
		var5=I%num%


		if part =             ; PART ������ null�̸�
			{
				MsgBox,�����մϴ�
				ExitApp
			}


		;MouseClick, left,  158,  244
		;Send, {SHIFTDOWN}{END}{SHIFTUP}{DEL}

		;Click,826,22
		Sleep 300
		MouseClickDrag,left,576,244,698,241
		Sleep 300
		Send {Del}
		Sleep 300
		Send %part%
		Sleep 300
		Send {Enter}

		Sleep 1500
		Clipboard=
        Click 128, 408, 2
		sleep 300
        send ^c
		Sleep 300
        ClipWait, 0

		Clipboard := Clipboard  ; ����� ����, HTML, �Ǵ� ��Ÿ ������ �ؽ�Ʈ�� ����� �ؽ�Ʈ�� ��ȯ�մϴ�.

		if (Clipboard ="SCQ")
			X1.Range(var5).value := Clipboard

		;~ if  Clipboard = SCQ
		;~ {
			;~ Click,553,23
			;~ Sleep 300
			;~ Click,1697,210
			;~ Sleep 300
			;~ Click,662,336
			;~ Sleep 300
			;~ Send %part%
			;~ Sleep 300
			;~ Send {Tab}
			;~ Sleep 300
			;~ Send %qty%
			;~ Sleep 300
			;~ Click,1821,209
			;~ Send 300
			;~ Send {Enter}
			;~ X1.Range(var5).value := OK
		;~ }

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