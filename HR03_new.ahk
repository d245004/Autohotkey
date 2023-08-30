;~ #SingleInstance Force
;~ #NoEnv
;~ SetWorkingDir %A_ScriptDir%
;~ SetBatchLines -1

; HR03 청구 수량 조정
^F2::
    Loop 1000
    {
        aa = 1
        num = 384
        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1
        while (tmp1)
        {
            Click,1390,%num%,2
            Sleep, 500
            Send, ^c
            Sleep, 300
            ClipWait, 0
            if (Clipboard = "")
            {
                MsgBox, 데이터가 없는 모양인데 작업을 끝내야겠네
                eend = 0
                Reload
            }

            if (Clipboard = 0)
            {
                Click,1273,%num%,2
                sleep,500
                Send,^c
                sleep,300

                Click,880,%num%
                Sleep,300
                send,{Enter}
                sleep,300
                send,^v
                sleep,300

                Click,1321,311
                Sleep,1000

                click,845,177

                sleep,1500
                num -= 45
                aa -= 1
                Continue
            }
            num += 45
            aa += 1
            if (aa>7)
                Break
        }
        if (eend != 0)
        {
            Click,1451,311
            Sleep, 1000
            Continue
        }
        Break
    }
return

^Space::Pause

^PGDN::ExitApp

^PgUp::Reload