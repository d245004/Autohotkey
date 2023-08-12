SendMode Input   ;  send 명령을 빠르게 실행하기위해서 사용
;CoordMode, Mouse, Screen

num=2
XLS_file_path3 := A_WorkingDir . "HP21.XLSX"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:H").NUMBERFORMAT := "@"   ; 엑셀 셀을 TEXT로 고정함


;부품 리스트 조회
F2::
InputBox,num,시작 행 선택,시작 행을 입력 하시요. 기본 시작은 2를 입력하시요.

Loop 1000
{	
	WinGetActiveTitle,tmp1
	while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
	{
		VAR=D%num%            ; ㅣLEP 지정
		VAR2=B%num%           ; PART  지정
		var3=G%num%
		var4=H%num%
		var5=I%num%
		var6=J%num%
		var7=K%num%
		part := X1.Range(VAR2).value
		if part =    ;변수 내용이 null이면
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
		X1.Range(var4).value := Clipboard   ; doprice D열
		
		Clipboard=
		Click,1109,377,2
		Sleep,500
		Send ^c
		clipwait,0
		X1.Range(var5).value := Clipboard   ; doprice D열

		Clipboard=
		Click,330,384,2
		Sleep,500
		Send ^c
		clipwait,0
		X1.Range(var3).value := Clipboard   ; doprice D열

		Clipboard=
		Click,385,404,2
		Sleep,500
		Send ^c
		clipwait,0
		X1.Range(var6).value := Clipboard   ; doprice D열

		Clipboard=
		Click,715,376,2
		Sleep,500
		Send ^c
		clipwait,0
		X1.Range(var7).value := Clipboard   ; doprice D열
		
		num++
	}

	if part =
	{
		MsgBox,프로그램을 종료합니다.
		ExitApp
	}
}



return

PGUP::Pause    ; SPACE 키는 쓰면 안된다
	
PGDN::ExitApp