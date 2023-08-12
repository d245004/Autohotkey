SendMode Input   ;  send 명령을 빠르게 실행하기위해서 사용

XLS_file_path3 := A_WorkingDir . "VAT.XLSX"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:H").NUMBERFORMAT := "@"   ; 엑셀 셀을 TEXT로 고정함
return

F2::
InputBox,num,시작번호 입력,시작 번호를 입력하세요. (기본은 2로 입력)
if ErrorLevel
	ExitApp

Loop 
{	
Loop 
{
	WinGetActiveTitle,tmp1
		while (tmp1 = "HYUNDAI MOBIS_SUB 코드 조회 - Chrome")
	{
		var1=B%num%   ;  B열  code1
		var2=C%num%   ;  C열  code2
		var3=A%num%   ;  A열  차량번호
		car := X1.Range(VAR3).value
		Click,943,72
		Sleep,1000
		Click,785,108
		Sleep,1000
		Click,351,110
		Send,%car%
		if car =    ;변수 내용이 null이면
			{
				MsgBox, 131120, 작업 중지 여부, 자료 없음. 종료합니다
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

PGUP::Pause   ; SPACE 키는 쓰면 안된다
	
		

PGDN::ExitApp