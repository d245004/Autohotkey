SendMode Input   ;  send ����� ������ �����ϱ����ؼ� ���
;CoordMode, Mouse, Screen

num=2
XLS_file_path3 := A_WorkingDir . "HP21.XLSX"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:H").NUMBERFORMAT := "@"   ; ���� ���� TEXT�� ������


;��ǰ ����Ʈ ��ȸ
F2::
InputBox,num,���� �� ����,���� ���� �Է� �Ͻÿ�. �⺻ ������ 2�� �Է��Ͻÿ�.

Loop 1000
{	
	WinGetActiveTitle,tmp1
	while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
	{
		VAR=D%num%            ; ��LEP ����
		VAR2=B%num%           ; PART  ����
		var3=G%num%
		var4=H%num%
		var5=I%num%
		var6=J%num%
		var7=K%num%
		part := X1.Range(VAR2).value
		if part =    ;���� ������ null�̸�
			{
				break
			}
		lep := X1.Range(VAR).value
		
		Click,198,235
		Sleep,300
		if (lep = "H")
			Click,196,260
		
		if (lep = "K")
			Click,196,280
		Click,324,236,2
		Sleep,300
		Send {del}
		Sleep,300
		Send %part%
		Sleep,300
		Click,1443,200
		Sleep,500
		
		Clipboard=
		Click,1109,400,2
		Sleep,500
		Send ^c
		clipwait,0
		X1.Range(var4).value := Clipboard   ; doprice D��
		
		Clipboard=
		Click,1109,377,2
		Sleep,500
		Send ^c
		clipwait,0
		X1.Range(var5).value := Clipboard   ; doprice D��

		Clipboard=
		Click,330,384,2
		Sleep,500
		Send ^c
		clipwait,0
		X1.Range(var3).value := Clipboard   ; doprice D��

		Clipboard=
		Click,385,404,2
		Sleep,500
		Send ^c
		clipwait,0
		X1.Range(var6).value := Clipboard   ; doprice D��

		Clipboard=
		Click,715,376,2
		Sleep,500
		Send ^c
		clipwait,0
		X1.Range(var7).value := Clipboard   ; doprice D��
		
		num++
	}

	if part =
	{
		MsgBox,���α׷��� �����մϴ�.
		ExitApp
	}
}



return

PGUP::Pause    ; SPACE Ű�� ���� �ȵȴ�
	
PGDN::ExitApp