CoordMode, Mouse,Caret, Screen


; XLS_file_path3 := A_WorkingDir . "1013�������-����Ŀ�ӽ� ����10��22��.xlsx"
; Xl := ComObjCreate("Excel.Application")
; Xl := ComObjActive("Excel.Application")
; Xl.Range("A:H").NUMBERFORMAT := "@"   ; ���� ���� TEXT�� ������

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True


; ��ǰ�� û���Է� �� ��ġ
^F2::
    InputBox,num,���� �� ����,���� ���� �Է� �Ͻÿ� �⺻ ������ 4�� �Է��Ͻÿ�.
    InputBox,code,CODE ����,�븮�� �ڵ带 �Է� �Ͻÿ� (2450 or A041)
    InputBox, Npart, part column choice, part column �Է��Ͻÿ� (B or F ���)

    Loop
    {	
        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
        while (tmp1)
        {
            VAR2=%Npart%%num% ;    B���� �� ��°   part number
            part := Xl.Range(VAR2).value
            if (part="")
            {
                MsgBox,���α׷� ����
                ExitApp
            }
            var6=%Npart%%num% ;  PART 
            ; var10=D%num%  ;  D��  doprice
            var3=F%num% ;  F��  ��� ������� ����
            var4=H%num% ;  H��  ����� �������
            var7=G%num% ;  G��  ��� ������
            var8=I%num% ;  I��  ����� ������
            var5=J%num% ;  J��  ���������� ����
            var9=K%num% ;  K��  ��������� ���� 
            Click,181,264
            send %part%
            Sleep,300
            send {enter}
            sleep 3000 ;   1000�� 1�ʸ� �ǹ��Ѵ�

            Clipboard=
            Click,841,395,2
            Sleep,500
            Send ^c
            clipwait,0
            if (code = 2450)
                Xl.Range(var3).value := Clipboard ; ���������� ���� F��
            else
                Xl.Range(var7).value := Clipboard ; ��񽺱����� ���� G��

            Clipboard=
            Click,1210,451,2
            Sleep,500
            send ^c
            clipwait,0	
            if (code = 2450)
                Xl.Range(var4).value := clipboard ;  ������������  H��
            else 
                Xl.Range(var8).value := clipboard ;  ����ݱ�����  I��

            Clipboard=
            ;MouseClickDrag,L,538,295,640,295
            Click,1203,294,2
            Sleep,500
            Send ^c
            clipwait,0
            if (code = 2450)
                Xl.range(var5).value := Clipboard ;   ���������� ���� J��
            else
                Xl.range(var9).value := Clipboard ;   ��������� ���� K��

            Click,1100,202
            Sleep,200
            num++		
            sleep 2000
            ;Pause
        }
    }
Return

^Space::Pause

^PGUP::Reload

^PGDN::ExitApp ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�