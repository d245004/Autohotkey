#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True

^F2::
InputBox,num,시작 행 선택,시작 행을 입력 하시요. `n기본 시작은 2를 입력하시요.


Loop 1000
{
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1
	while (tmp1)
	{
		VAR2=E%num%
		VAR6=F%num%
		part := Xl.Range(VAR2).value

		if (part = "" )
		{
			; MsgBox,작업을 종료 합니다
			; ExitApp
			MsgBox, 계속 작업 하려면 F2 키를 누르시요
			Return
		}
		Click,384,113,2
		Sleep,500
		send %part%
		send {enter}
		sleep 1500     ;   1000은 1초를 의미한다

		Clipboard=
            MouseClickDrag,L,88,234,168,240 ;ENGLISH
            Sleep,500
            Send ^c
            clipwait,0
            Xl.range(var6).value := Clipboard



		Sleep 1500
		num++
		;Pause   ;잠시 중지
	}
}


return    ; 프로그램 원위치로


^Space::Pause   ; 누르면 잠시 중지를 ON / OFF 역활을 한다

^PGDN::ExitApp   ; 프로그램을 종료한다, 메모리에서 삭제된다

^PgUp::Reload