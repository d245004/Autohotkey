#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True

Gui Add, Button, x8 y184 w237 h23 gBTN, START
Gui Add, Edit, x104 y16 w27 h21 vNum,2
Gui Add, Edit, x104 y49 w27 h21 vM_part,A
Gui Add, Edit, x104 y82 w27 h21 vM_qty,B
Gui Add, Text, x16 y16 w81 h23 +0x200, 시작 번호
Gui Add, Text, x16 y48 w81 h23 +0x200, 부품번호 열
Gui Add, Text, x16 y80 w81 h23 +0x200, 수량 열

Gui Show, w253 h216, Window
Return

GuiEscape:
GuiClose:
    ExitApp

BTN:
;HR07 지원센타 청구입력
    Gui, Submit, NoHide
    WinActivate ahk_class Chrome_WidgetWin_1
    send +^{NumpadAdd}

    Loop 1000
    {
        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1
        while (tmp1)
        {
            VAR2=%M_part%%num%
            var3=%M_qty%%num%
            var4=U%num%
            part := Xl.Range(VAR2).value
            qty := round((Xl.range(var3).value),0)
            if (part = "" )
            {
                MsgBox, 계속 작업 하려면 F2 키를 누르시요
                Return
            }
            Click,1374,203
            Sleep,500
            Click,164,240
            Sleep 500
            send %part%
            send {enter}
            sleep 1500 ;   1000은 1초를 의미한다
            ;~ Click,529,369,2
            ;~ Sleep, 500
            send %qty%
            Sleep,500
            Clipboard=
            Click,258,376,2
            send ^c
            ClipWait, 0
            if (clipboard = "N")
            {
                Xl.Range(VAR4).value := "재고 없다"
                Pause
                num++
                Continue
            }
            Sleep,500
            Click,652,376
            Sleep, 1000
            Xl.Range(VAR4).value := "OK"
            num++
            Pause
        }
    }
return

^Space::Pause

^PGDN::ExitApp

^PgUp::Reload