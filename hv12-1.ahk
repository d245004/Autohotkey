;Coordmode, Mouse, Screen   ;  전체화면에서 마우스 포인트를 사용한다고 설정하는 것
num=2
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
		VAR3=E%num%
		QTY := round(X1.range(var3).value,0)
		var5=J%num%


		if part =             ; PART 내용이 null이면
			{
				MsgBox,종료합니다
				ExitApp
			}
			Click,1261,207
			Sleep 1300
			Click,209,336
			Sleep 1300
			Send %part%
			Sleep 1300
			Send {Tab}
			Sleep 1300
			Send %qty%
			Sleep 1300
			Click,1385,206
			Send 1300
			Send {Enter}
			Sleep 1000
			X1.Range(var5).value := "OK"

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