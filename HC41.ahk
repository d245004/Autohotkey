SendMode Input ;  send ����� ������ �����ϱ����ؼ� ���
Coordmode, Mouse, Screen ;  ��üȭ�鿡�� ���콺 ����Ʈ�� ����Ѵٰ� �����ϴ� ��

num=2
XLS_file_path3 := A_WorkingDir . "HC41.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
;X1.Range("A:F").NUMBERFORMAT := "@"

; �� û���� �� ���α׷�

^F2:: 
    Loop 1000
    {	
        ;WinGetActiveTitle,tmp1
        ;	while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
        while (tmp1)

        {
            VAR2=B%num% ;    B���� �� ��°   CODE
            code := X1.Range(VAR2).value
            if (code = "")
            {
                MsgBox,�۾� �Ϸ�! (�����߳׿�.)
                ExitApp
            }
            var4=E%num% ;  F��  û����
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

            sleep 10000 ;   1000�� 1�ʸ� �ǹ��Ѵ�

            Clipboard=
            MouseClickDrag,L,664,297,730,297 ; ���콺 �巡��
            Send ^c
            ClipWait,0
            X1.range(var4).value := Clipboard ; F�� û����

            num++		
            sleep 1000
        }
        Pause ;��� ����
    }

return ; ���α׷� ����ġ��

^PGUP::Reload 

^PGDN::ExitApp ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�

^Space::Pause
