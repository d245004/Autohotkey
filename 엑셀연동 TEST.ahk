num=3
XLS_file_path3 := A_WorkingDir . "\hv01.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
;����.range("A:D").NUMBERFORMAT := "@"
X1.Range("A:E").NUMBERFORMAT := "@"

;SetFormat,FLOAT,1,1
return

^F1::
Loop 1000      ; 1000���� ���ƶ�
{
WinGetActiveTitle,tmp1
	
while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome") 
{
VAR=D%num%
;VAR1=C%num%
var2=B%num%
as2 := X1.Range(VAR2).VALUE
AS := ROUND(X1.Range(VAR).VALUE,0)  ;  �Ҽ� 2�ڸ� �ȳ�����
AS1 := H

; PART NUMBER �ִ� CELL��  =TEXT(��,"0")  ����  ��ȯ�� ���� �۾� �� ��



Send %AS1%   
Sleep 1000    ; 1000  ->  1��
;MsgBox PART %AS2%
Send {TAB}    ; ONE
;if NUM > 10
;	SEND {SPACE}  ; ������ K�� �з½� ���� �߻��Ҷ� ��� , HAIMS�� BUG ��
send %as2%   
sleep 1000
send {tab}   ; TWO
Sleep 1000

;Send {TAB}
send %AS%
;sleep 1000
send {tab} 
;send %AS1%
;Sleep 1000	
send {tab} 

num++

;if (num=24)
	Pause
;if (num=44)
;	Pause

break
}
}
;X1.QUIT
return                 ;����ġ�� ���� �϶�� 

Space::Pause           ; pause key �� space �� �ϸ� �ȵȴ� (������ �߻��ؼ� ����� �۵� ����)

End::ExitApp