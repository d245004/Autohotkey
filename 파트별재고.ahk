;Coordmode, Mouse, Screen   ;  전체화면에서 마우스 포인트를 사용한다고 설정하는 것
num=2
XLS_file_path3 := A_WorkingDir . "HR07.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:E").NUMBERFORMAT := "@"
return

F2::
Loop 1000
{	
	WinGetActiveTitle,tmp1
		while (tmp1 = "파트별재고 - Whale")
	{
		VAR2=C%num%   ;    C열의 몇 번째
		var3=E%num%
		part := X1.Range(VAR2).value
		MouseClick, left,108,146
		Send, {SHIFTDOWN}{END}{SHIFTUP}{DEL}
		send %part%
		send {enter}
		sleep 1000     ;   1000은 1초를 의미한다
		
		if
			MouseClick,L,215,400   ;  품목이 H 일때
		else
			MouseClick,L,215,420   ;  품목이 K 일때
		Sleep,1000
		MouseClick,L,208,361
		Sleep,1000
		MouseClick,L,271,691
		Sleep,1000
		WinGetActiveTitle, tmp2
			while (tmp2 = "재고조정 - Whale")
			{
				qty := round(X1.range(var3).value,0)
		        send %qty%
				send {enter}
				Sleep,1000
				send {enter}
				Sleep,1000
				MouseClick,L,219,426
				Exit
			}
		num++
		Pause   ;잠시 중지
}
Pause  ;잠시 중지
}

return    ; 프로그램 원위치로

PGUP::Pause   ; 누르면 잠시 중지를 ON / OFF 역활을 한다

PGDN::ExitApp   ; 프로그램을 종료한다, 메모리에서 삭제된다