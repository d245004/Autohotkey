#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
; CoordMode, Mouse,Caret, Screen,Pixel


; XLS_file_path3 := A_WorkingDir . "11월미래조치건.xls"
; Xl := ComObjCreate("Excel.Application")
; Xl := ComObjActive("Excel.Application")

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True
; Xl 과 Xl 구분 할 것 (에러 발생 이유)

; 전문점 수령으로 주문하는 program

^F2::
InputBox,num,시작 행 선택,시작 행을 입력 하시요. `n기본 시작은 2를 입력하시요.
InputBox,pa,part 열 선택,part열을 입력
InputBox,qt,수량 열 선택,qty 열 입력

Loop 1000
{
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1
	while (tmp1)
	{
		VAR2=%pa%%num%
		var3=%qt%%num%
		var4=K%num%
		part := Xl.Range(VAR2).value
		qty := round((Xl.range(var3).value),0)

		if (qty<1)
        {
            num++
            Continue
            Xl.Range(VAR4).value := "OK"
        }
		if (part = "" )
		{
			; MsgBox,작업을 종료 합니다
			; ExitApp
			MsgBox, 계속 작업 하려면 F2 키를 누르시요
			Return
		}
		Click,1374,203
		Sleep,500
		Click,164,240
		Sleep 500

		send %part%
		send {enter}
		sleep 1500     ;   1000은 1초를 의미한다
		send %qty%
		Sleep,1500
		Click,566,370
		Sleep,1000
		Send {DOWN} {ENTER}
		Sleep.1000
		Click,668,370
		Sleep, 1000

		;~ ImageSearch, ax, ay, 0,0,1387,186,*transFF977D *60 std.png
		;~ if ErrorLevel = 0
		;~ {
			;~ ; MsgBox, 찾았다
			;~ Send, {enter}
			;~ Xl.Range(var4).value := "OK"
		;~ }

		;~ if ErrorLevel = 1
		;~ {
			;~ ; MsgBox, 못찾는데..
			;~ Send, {right}
			;~ Send, {enter}
			;~ Xl.Range(var4).value := "재고 없네"
		;~ }

		Xl.Range(VAR4).value := "OK"
		num++
		Pause   ;잠시 중지
	}
}


return    ; 프로그램 원위치로

^Right::
	Send, {right}
	Send, {enter}
	Xl.Range(var4).value := "재고 없네"

^Down::
	Send, {right}
	Send, {enter}
	pause

^Space::Pause   ; 누르면 잠시 중지를 ON / OFF 역활을 한다

^PGDN::ExitApp   ; 프로그램을 종료한다, 메모리에서 삭제된다

^PgUp::Reload