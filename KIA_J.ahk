#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

; CoordMode, Mouse,Caret, Screen,Pixel

; XLS_file_path3 := A_WorkingDir . "��ǰ�� ���� ����(20211027).xls"
; Xl := ComObjCreate("Excel.Application")
; Xl := ComObjActive("Excel.Application")

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True
; Xl �� Xl ���� �� �� (���� �߻� ����)

^F2::
    InputBox,num,���� �� ����,���� ���� �Է� �Ͻÿ�. �⺻ ������ 3�� �Է��Ͻÿ�.
    InputBox,s_lep,LEP ����,LEP ���� �Է� �Ͻÿ�.
    InputBox,s_part,PART �� ����,PART ���� �Է� �Ͻÿ�.

    aa = 1
    Loop
    {	
        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
        while (tmp1)
        {
            VAR2=%s_part%%num% ;   3143029200 A���� �� ��°   part number
            part := Xl.Range(VAR2).value
            if (part="")
            {
                MsgBox,Press any key
                return
            }
            var6=H%num% ;  E��  PART NAME
            var10=I%num% ;  D��  doprice
            var11=J%num%
            var17=%s_lep%%num% ;  C��  lep
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
            Click,235,600,2
            Sleep,500
            Send ^c
            clipwait,0
            Xl.Range(var10).value := Clipboard ; doprice D��

            Clipboard=
            MouseClickDrag,L,416,336,607,336 ;ENGLISH
            ;~ MouseClickDrag,L,710,336,910,336    ;KOREAN
            Sleep,500
            Send ^c
            clipwait,0
            Xl.range(var6).value := Clipboard ;   partname E��

            Clipboard=
            MouseClickDrag, L, 309, 909, 421, 909
            Sleep,500
            Send ^c
            clipwait,0
            Xl.Range(var11).value := Clipboard ; �����̼�


            Sleep,300
            num++		
            sleep 1000
            aa += 1
            if (aa > 1300)
            {
                aa = 1
                Pause
            }
        }
    }

Return

^Space::Pause ; ������ ��� ������ ON / OFF ��Ȱ�� �Ѵ�

^PgUp::Reload

^PGDN::ExitApp ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�
