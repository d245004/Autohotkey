SendMode Input   ;  send ����� ������ �����ϱ����ؼ� ���

XLS_file_path3 := A_WorkingDir . "VAT.XLSX"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:H").NUMBERFORMAT := "@"   ; ���� ���� TEXT�� ������
return

F2::
InputBox,num,���۹�ȣ �Է�,���� ��ȣ�� �Է��ϼ���. (�⺻�� 2�� �Է�)
if ErrorLevel
	ExitApp

Loop 
{	
Loop 
{
	WinGetActiveTitle,tmp1
		while (tmp1 = "HYUNDAI MOBIS_SUB �ڵ� ��ȸ - Chrome")
	{
		var1=B%num%   ;  B��  code1
		var2=C%num%   ;  C��  code2
		var3=A%num%   ;  A��  ������ȣ
		car := X1.Range(VAR3).value
		Click,943,72
		Sleep,1000
		Click,785,108
		Sleep,1000
		Click,351,110
		Send,%car%
		if car =    ;���� ������ null�̸�
			{
				MsgBox, 131120, �۾� ���� ����, �ڷ� ����. �����մϴ�
				ExitApp
			}
		Sleep,1000
		Click,1002,74
		Sleep,5000
		Clipboard=
		MouseClickDrag,L,83,204,170,204
		Sleep,1500
		Send ^c
		clipwait,0
		X1.range(var1).value := Clipboard
		num++
		break
	}
}
}
return

PGUP::Pause   ; SPACE Ű�� ���� �ȵȴ�
	
		

PGDN::ExitApp