#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

; Coordmode, Mouse, Screen   ;  ��üȭ�鿡�� ���콺 ����1Ʈ�� ����Ѵٰ� �����ϴ� ��
; XLS_file_path3 := A_WorkingDir . "11��5�� ���.xlsx"
; Xl := ComObjActive("Excel.Application")
; Xl.Range("A:C").NUMBERFORMAT := "@"

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True

^F2::
    InputBox,num,���� �� ����,���� ���� �Է� �Ͻÿ�. �⺻ ������ 3�� �Է��Ͻÿ�.
    InputBox,M_part,,part number ���� �Է� �Ͻÿ�
    InputBox,M_lep,,LEP ���� �Է� �Ͻÿ�
    InputBox,count_aa,,�ݺ��� Ƚ���� �Է��Ͻÿ�

    aa = 1
    Loop
    {	
        ;WinGetActiveTitle,tmp1
        ;while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")

        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
        while (tmp1)
        {
            VAR2=%M_part%%num% ;   3143029200 A���� �� ��°   part number
            part := Xl.Range(VAR2).value
            if (part="")
            {
                MsgBox,Press any key
                return
            }
            var6=D%num% ;  D�� ��ǰ��
            var10=F%num% ;  E��  �����̼�
            var17=%M_lep%%num% ;  C��  lep
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
            sleep 1500 ;   1000�� 1�ʸ� �ǹ��Ѵ�

            Clipboard=
            MouseClickDrag, L, 309, 909, 421, 909
            Sleep,500
            Send ^c
            clipwait,0
            Xl.Range(var10).value := Clipboard ; �����̼�

            Clipboard=
            MouseClickDrag,L,416,336,607,336    ;ENGLISH
            ; MouseClickDrag,L,710,336,910,336 ;KOREAN
            Sleep,500
            Send ^c
            clipwait,0
            Xl.range(var6).value := Clipboard ;   partname E��

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

^Space::Pause ; ������ ��� ������ ON / OFF ��Ȱ�� �Ѵ�

^PGDN::ExitApp ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�

^PgUp::Reload