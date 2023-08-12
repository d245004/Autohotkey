SendMode Input ;  send 명령을 빠르게 실행하기위해서 사용
Coordmode, Mouse, Screen ;  전체화면에서 마우스 포인트를 사용한다고 설정하는 것

num=2
XLS_file_path3 := A_WorkingDir . "HC41.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
;X1.Range("A:F").NUMBERFORMAT := "@"

; 월 청구액 비교 프로그램

^F2:: 
    Loop 1000
    {	
        ;WinGetActiveTitle,tmp1
        ;	while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
        while (tmp1)

        {
            VAR2=B%num% ;    B열의 몇 번째   CODE
            code := X1.Range(VAR2).value
            if (code = "")
            {
                MsgBox,작업 완료! (지루했네요.)
                ExitApp
            }
            var4=E%num% ;  F열  청구액
            Click,600,238
            Send,{DEL}
            Sleep,300
            send %code%
            sleep 500
            Send {down}
            Sleep 500
            Send {ENTER}
            sleep 500
            send {enter}

            sleep 10000 ;   1000은 1초를 의미한다

            Clipboard=
            MouseClickDrag,L,664,297,730,297 ; 마우스 드래그
            Send ^c
            ClipWait,0
            X1.range(var4).value := Clipboard ; F열 청구액

            num++		
            sleep 1000
        }
        Pause ;잠시 중지
    }

return ; 프로그램 원위치로

^PGUP::Reload 

^PGDN::ExitApp ; 프로그램을 종료한다, 메모리에서 삭제된다

^Space::Pause
