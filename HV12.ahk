;Coordmode, Mouse, Screen   ;  전체화면에서 마우스 포인트를 사용한다고 설정하는 것
num=110
XLS_file_path3 := A_WorkingDir . "보문 재고 익스포트.xlsx"
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
		VAR2=B%num%   ;    B열의 몇 번째
		part := X1.Range(VAR2).value
		VAR3=I%num%
		QTY := round(X1.range(var3).value,0)
		var5=I%num%


		if part =             ; PART 내용이 null이면
			{
				MsgBox,종료합니다
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

		Clipboard := Clipboard  ; 복사된 파일, HTML, 또는 기타 형식의 텍스트를 평범한 텍스트로 변환합니다.

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
		;~ Pause   ;잠시 중지
}
Pause  ;잠시 중지
}

return    ; 프로그램 원위치로

^PGUP::Reload

^Space::Pause

^PGDN::ExitApp   ; 르로그램을 종료한다, 메모리에서 삭제된다