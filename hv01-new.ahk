#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
CoordMode, Mouse, Screen

num=2
XLS_file_path3 := A_WorkingDir . "DELETE.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
;X1.Range("A:D").NUMBERFORMAT := "@"   ; ���� ���� TEXT�� ������
return


;HV01  ���� �Է� ����
^F2::
Loop 1000
{
	InputBox,num,,Start Number Input!
	InputBox,aa,�Է� �Ǽ���?
		Click,1672,210
		Sleep 500
;		Click,1131,238
;		Sleep 200
;		Send {Down}
;		Sleep 200
;		Send {Down}
;		Sleep 200
;		Send {Enter}
;		Sleep 500
		Click,600,589
		Sleep 500
		Send "3�� ���̵� ��� ���"
		Sleep 300

		Click,570,335,2
		Sleep 1000
		Send {DELETE}

aaa = 1
Loop %aa%
{
;	WinGetActiveTitle,tmp1
;		while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
		tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1
		while (tmp1)

{
		VAR = B%num%            ; A�� ����   LEP
		VAR2 = C%num%           ; B�� ����   PART NUMBER
		var4 = E%num%           ; C�� ����   QTY
		var5 = G%num%           ; D�� ����   "�Է� �Ϸ�"
		part := X1.Range(VAR2).value
;		lep := substr(part,1,1)  ;�������� �����Ѻκ� �ؽ�Ʈ�� �̾Ƴ�
;		part := substr(part,2,15)  ;;�������� �����Ѻκ� �ؽ�Ʈ�� �̾Ƴ�
		;MsgBox,%part%
		if part =             ; PART ������ null�̸�
			{
				MsgBox,�����մϴ�
				ExitApp
			}
		qty := round(X1.range(var4).value,0)
		lep := X1.Range(VAR).value

		send %lep%		;  LEP ����
		send {tab}      ;  PART�� �̵�
		send %part%     ;  PART �Է�
		sleep 2000
		send {tab}      ;  â���� �̵�
		sleep 1400
		send {tab}		;  ���� �������� �̵�
		send -%qty%		;  ���� ���� �Է�
		Sleep 1500
;		sleep 3500       ;  2.5�� ���   (���̸� ERROR �߻�)
		send {tab}		;  �����ܰ��� �̵�
		sleep 1500      ;  1�� ���
		Send {tab}      ;  ���� �� LEP�� �̵�
		abc = %aaa%
		aaa++
		X1.Range(var5).value := abc "  �Է� �Ϸ�"

		num++
	    break
	}
}
Click,1797,210
;Sleep 20000
Pause
;Click,1672,210
;Sleep 500
}

return
^Space::Pause

^PGUP::Reload

^PGDN::ExitApp