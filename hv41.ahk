#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
;Coordmode, Mouse, Screen   ;  ��üȭ�鿡�� ���콺 ����Ʈ�� ����Ѵٰ� �����ϴ� ��

; XLS_file_path3 := A_WorkingDir . "�̻��������(10��).xlsx"
; X1 := ComObjActive("Excel.Application")
; X1.Range("A:K").NUMBERFORMAT := "@"

FileSelectFile, path
X1 := ComObjCreate("Excel.Application")
X1.Workbooks.Open(path)
X1.Visible := True


^F2::
    InputBox,num,���� �� ����,�����Ϸ��� �⺻ ���� 2�� �Է��Ͻÿ�

    aa = 1
    Loop
    {
        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1
        while (tmp1)
        {
            VAR2=E%num% ;   A���� �� ��°   part number
            part := X1.Range(VAR2).value
            if (part="")
            {
                MsgBox,Press any key
                return
            }
            var4=I%num% ;���Լ���
            maip := X1.Range(var4).value
            var5=J%num% ;�Ǹż���
            machul := X1.Range(var5).value
            var6=K%num% ;�����˰��

            maip := Round(maip)
            machul := Round(machul)

            Click,193,263,2
            Sleep,500
            Send, {Home}
            Sleep, 500
            Send {Delete 20}
            Sleep,500
            send %part%
            Sleep,500
            Send, {Enter}
            sleep 1500 ;   1000�� 1�ʸ� �ǹ�

            Clipboard=
            Click,1205,390,2
            Send, ^c
            ClipWait, 0
            qty := Clipboard ; �����
            sleep,300

            Clipboard=
            Click,144,424,2
            Send, ^c
            ClipWait, 0
            jqty := Clipboard ; �׸��
            sleep,300

            Clipboard=
            Click,273,423,2
            Send, ^c
            ClipWait, 0
            total := Clipboard ; �Ǹż���
            Sleep, ,300

            sale :=

            ms := 0

            aa := 1
            While aa<=(jqty)
            {
                Clipboard=
                MouseClickDrag,L,508,495+ms,656,495+ms
                Sleep,300
                Send ^c
                clipwait,0
                sale := sale Clipboard ","
                Sleep, 300

                ; Clipboard=
                ; MouseClick, L, 770, 495+ms,2
                ; Sleep, 300
                ; Send, ^c
                ; ClipWait, 0
                ; sale := sale Clipboard ","
                ; Sleep, 300

                aa += 1
                if (aa>13)
                    Break
                ms += 22
            }
            X1.range(var6).value := "�԰�" maip "��,�Ǹ�" total "��(" sale ")���" qty

            Sleep,300
            num++
            sleep 500
            aa += 1
            if (aa > 1300)
            {
                aa = 1
                Pause
            }
            ; Pause
            Sleep, 1000
        }
    }

Return

^Space::Pause

^PGUP::Reload

^PGDN::ExitApp ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�
