#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
; CoordMode, Mouse,Caret, Screen,Pixel


; XLS_file_path3 := A_WorkingDir . "11���̷���ġ��.xls"
; Xl := ComObjCreate("Excel.Application")
; Xl := ComObjActive("Excel.Application")

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True
; Xl �� Xl ���� �� �� (���� �߻� ����)

; ������ �������� �ֹ��ϴ� program

^F2::
InputBox,num,���� �� ����,���� ���� �Է� �Ͻÿ�. `n�⺻ ������ 2�� �Է��Ͻÿ�.
InputBox,pa,part �� ����,part���� �Է�
InputBox,qt,���� �� ����,qty �� �Է�

Loop 1000
{
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1
	while (tmp1)
	{
		VAR2=%pa%%num%
		var3=%qt%%num%
		var4=K%num%
		part := Xl.Range(VAR2).value
		qty := round((Xl.range(var3).value),0)

		if (qty<1)
        {
            num++
            Continue
            Xl.Range(VAR4).value := "OK"
        }
		if (part = "" )
		{
			; MsgBox,�۾��� ���� �մϴ�
			; ExitApp
			MsgBox, ��� �۾� �Ϸ��� F2 Ű�� �����ÿ�
			Return
		}
		Click,1374,203
		Sleep,500
		Click,164,240
		Sleep 500

		send %part%
		send {enter}
		sleep 1500     ;   1000�� 1�ʸ� �ǹ��Ѵ�
		send %qty%
		Sleep,1500
		Click,566,370
		Sleep,1000
		Send {DOWN} {ENTER}
		Sleep.1000
		Click,668,370
		Sleep, 1000

		;~ ImageSearch, ax, ay, 0,0,1387,186,*transFF977D *60 std.png
		;~ if ErrorLevel = 0
		;~ {
			;~ ; MsgBox, ã�Ҵ�
			;~ Send, {enter}
			;~ Xl.Range(var4).value := "OK"
		;~ }

		;~ if ErrorLevel = 1
		;~ {
			;~ ; MsgBox, ��ã�µ�..
			;~ Send, {right}
			;~ Send, {enter}
			;~ Xl.Range(var4).value := "��� ����"
		;~ }

		Xl.Range(VAR4).value := "OK"
		num++
		Pause   ;��� ����
	}
}


return    ; ���α׷� ����ġ��

^Right::
	Send, {right}
	Send, {enter}
	Xl.Range(var4).value := "��� ����"

^Down::
	Send, {right}
	Send, {enter}
	pause

^Space::Pause   ; ������ ��� ������ ON / OFF ��Ȱ�� �Ѵ�

^PGDN::ExitApp   ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�

^PgUp::Reload