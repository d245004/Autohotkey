CoordMode, Mouse,Caret, Screen
XLS_file_path3 := A_WorkingDir . "list_song.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:H").NUMBERFORMAT := "@"   ; 엑셀 셀을 TEXT로 고정함
return


;업체별 부품 가격 적용율 관리
F2::
InputBox,num,시작 행 선택,시작 행을 입력 하시요. 기본 시작은 2를 입력하시요.
Loop
{	
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
	while (tmp1)
	{
		var1=A%num%   ; code
		part := X1.Range(var1).value
		if (part="")
		{
			MsgBox,프로그램 종료
			ExitApp
		}

		Click,1260,200

		Click,100,320
		Sleep, 1200
		Send, %part%

		Click,100,340
		Sleep, 1200
		Send, %part%

		Click,100,365
		Sleep, 1200
		Send, %part%

		Click,100,390
		Sleep, 1200
		Send, %part%

		Click,405,320
		Sleep, 1200
		Send, II
		
		Click,405,340
		Sleep, 1200
		Send, IN
		
		Click,405,365
		Sleep, 1200
		Send, IP
		
		Click,405,390
		Sleep, 1200
		Send, TC
		
		
		Click,715,326
		Sleep, 1200
		Send, -5

		Click,715,340
		Sleep, 1200
		Send, -5

		Click,715,365
		Sleep, 1200
		Send, -5

		Click,715,390
		Sleep, 1200
		Send, -5


		Click,1397,198
		Sleep, 2000

		Click,846,172
		Sleep, 2000

		; Click,919,168
		; Sleep, 2000

		num++
		Pause
		
	}
Pause  ;잠시 중지
}

return    ; 프로그램 원위치로

PGUP::Pause   ; 누르면 잠시 중지를 ON / OFF 역활을 한다

PGDN::ExitApp   ; 르로그램을 종료한다, 메모리에서 삭제된다