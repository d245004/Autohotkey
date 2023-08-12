#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

num=2
XLS_file_path3 := A_WorkingDir . "HI01.XLS"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
;X1.Range("A:D").NUMBERFORMAT := "@"   ; 엑셀 셀을 TEXT로 고정함
return

;���԰� ���� ���α׷�
F2::
InputBox,num,시작 행 선택,시작하려는 행번호를 입력하시요

Loop
{	
    Click,1010,199
    Sleep, 300
    Click,745,294
    Sleep, 300
    Send, 0009
    Sleep, 300
    Send, {down} {enter}
    Sleep, 300
    Click,157,384,2
    Sleep, 300

    aa = 1
    Loop 
    {
		tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
		while (tmp1)
        {
		    var=A%num%            ; A열 지정   LEP
		    var2=B%num%           ; B열 지정   PART NUMBER
		    var4=C%num%           ; C열 지정   QTY
		    var5=D%num%           ; D열 지정   "입력 완료"
		    part := X1.Range(VAR2).value
		    if part =             ; PART 내용이 null이면
			{
				MsgBox,종료합니다
				ExitApp
			}
		    qty := round(X1.range(var4).value,0)
		    lep := X1.Range(VAR).value
            Send, {del}
            Sleep, 300
		    send %lep%		;  LEP 선택
            Sleep, 300
		    send {tab}      ;  PART로 이동
		    send %part%     ;  PART 입력
		    sleep 1000
		    send {tab}		;  조정 수량으로 이동
		    send %qty%		;  조장 수량 입력
		    Sleep 300
		    send {tab}		;  조정단가로 이동
		    X1.Range(var5).value := "--OK--"
		    sleep 500      ;  1초 대기
		    Send {tab}      ;  다음 줄 LEP로 이동

		    num++
            aa += 1
            if (aa>9)
                Break

	    }
        Break
    }
    Click,1130,198
    Pause
}

return

PGUP::
    Reload    ; SPACE 키는 쓰면 안된다
    Send, {F2}
^Space::Pause

PGDN::ExitApp