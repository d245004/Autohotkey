num=3
XLS_file_path3 := A_WorkingDir . "\hv01.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
;엑셀.range("A:D").NUMBERFORMAT := "@"
X1.Range("A:E").NUMBERFORMAT := "@"

;SetFormat,FLOAT,1,1
return

^F1::
Loop 1000      ; 1000번을 돌아라
{
WinGetActiveTitle,tmp1
	
while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome") 
{
VAR=D%num%
;VAR1=C%num%
var2=B%num%
as2 := X1.Range(VAR2).VALUE
AS := ROUND(X1.Range(VAR).VALUE,0)  ;  소수 2자리 안나오게
AS1 := H

; PART NUMBER 있는 CELL을  =TEXT(셀,"0")  으로  변환한 다음 작업 할 것



Send %AS1%   
Sleep 1000    ; 1000  ->  1초
;MsgBox PART %AS2%
Send {TAB}    ; ONE
;if NUM > 10
;	SEND {SPACE}  ; 차종을 K로 압력시 에러 발생할때 사용 , HAIMS의 BUG 임
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
return                 ;원위치로 복귀 하라는 

Space::Pause           ; pause key 를 space 로 하면 안된다 (간섭이 발생해서 제대로 작동 안함)

End::ExitApp