#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

; Coordmode, Mouse, Screen   ;  전체화면에서 마우스 포인1트를 사용한다고 설정하는 것
; XLS_file_path3 := A_WorkingDir . "11월5일 재고.xlsx"
; Xl := ComObjActive("Excel.Application")
; Xl.Range("A:C").NUMBERFORMAT := "@"

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True

^F2::
    InputBox,num,시작 행 선택,시작 행을 입력 하시요. 기본 시작은 3을 입력하시요.
    InputBox,M_part,,part number 열을 입력 하시요
    InputBox,M_lep,,LEP 열을 입력 하시요
    InputBox,count_aa,,반복할 횟수를 입력하시요

    aa = 1
    Loop
    {	
        ;WinGetActiveTitle,tmp1
        ;while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")

        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
        while (tmp1)
        {
            VAR2=%M_part%%num% ;   3143029200 A열의 몇 번째   part number
            part := Xl.Range(VAR2).value
            if (part="")
            {
                MsgBox,Press any key
                return
            }
            var6=D%num% ;  D열 부품명
            var10=F%num% ;  E열  로케이션
            var17=%M_lep%%num% ;  C열  lep
            gubunlep := Xl.Range(var17).value

            Click,985,202
            Sleep,300
            Click,226,239
            Sleep,300
            if (gubunlep = "H")
            {
                Click,180,260
            }
            if (gubunlep = "K")
            {
                Click,180,280
            }
            ;MsgBox,%gubunlep%
            Sleep,300
            Click,253,239
            Sleep,300
            send %part%
            Sleep,300
            Click,1046,202
            sleep 1500 ;   1000은 1초를 의미한다

            Clipboard=
            MouseClickDrag, L, 309, 909, 421, 909
            Sleep,500
            Send ^c
            clipwait,0
            Xl.Range(var10).value := Clipboard ; 로케이션

            Clipboard=
            MouseClickDrag,L,416,336,607,336    ;ENGLISH
            ; MouseClickDrag,L,710,336,910,336 ;KOREAN
            Sleep,500
            Send ^c
            clipwait,0
            Xl.range(var6).value := Clipboard ;   partname E열

            Sleep,300
            num++		
            sleep 1000
            aa += 1
            if (aa > count_aa)
            {
                aa = 1
                Pause
                Continue
            }
        }
    }
Return

^Space::Pause ; 누르면 잠시 중지를 ON / OFF 역활을 한다

^PGDN::ExitApp ; 프로그램을 종료한다, 메모리에서 삭제된다

^PgUp::Reload