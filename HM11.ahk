SendMode Input             ;  send ����� ������ �����ϱ����ؼ� ���
;Coordmode, Mouse, Screen   ;  ��üȭ�鿡�� ���콺 ����Ʈ�� ����Ѵٰ� �����ϴ� ��
num=2
XLS_file_path3 := A_WorkingDir . "MACHUL.xls"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:P").NUMBERFORMAT := "@"


; HAIMS ���ݰ�꼭  ���α׷�
F2:: 

InputBox,nal,���� ���� �Է�,���� ���ڸ� �Է� �ϼ���. ex-> 20190331

Loop 1000
{	
	WinGetActiveTitle,tmp1
		while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
	{
		VAR2=C%num%   ;    C���� �� ��°   CODE
		var3=M%num%   ;    �����ݾ�
		var4=N%num%   ;    �ΰ��� 
		code := X1.Range(VAR2).value
		price := round(X1.Range(var3).value,0)
		vat := round(X1.Range(var4).value,0)
		bigo := price + vat
		if (code = "")
			{
				MsgBox,�۾� �Ϸ�! (�����߳׿�.)
				ExitApp
			}
		Click,969,203
		Sleep,300
		Click,406,266,2
		Sleep,300
		Send,{del}
		Sleep,300
		Send,%nal%
		Sleep,300
		Click,245,371,2
		Sleep,300
		Send,{del}
		Sleep,300
		Send,%code%
		sleep 1300
		Send {down}
		send {enter}
		sleep 1300     
		Click,828,627,2
		Sleep,300
		Send,{del}
		Sleep,300
		Send,%price%
		Sleep,300
		Send,{tab}
		Sleep,300
		;  Click,987,625
		Send,%price%
		sleep,300
		Send,{tab}
		Sleep,300
		send,%vat%
		Sleep,300
		Send,{enter}
		Sleep,300
		;MouseMove,1177,779
		;Sleep,300	
		Click,1177,779,2
		Sleep,300
		Send,{del}
		Sleep,300
		send,%bigo%
		Sleep,300
		Send,{enter}
		Sleep,300
		MouseMove,1100,202
		Loop
			{
				getkeystate,vvar,enter
				IF vvar=D
				{
					Click,1100,202
					Sleep,300
					Send,{enter}
					; Click,842,174
					Sleep,3000
					Send,{enter}
					; Click,917,173
					Sleep,1000
					break
				}
			}
			
		num++	
		Pause
}
Pause  ;��� ����
}

return    ; ���α׷� ����ġ��

PGUP::Pause   ; ������ ��� ������ ON / OFF ��Ȱ�� �Ѵ�

PGDN::ExitApp   ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�