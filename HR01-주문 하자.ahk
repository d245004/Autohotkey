CoordMode, Mouse,Caret, Screen

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True


; ��ǰ�� û���Է� �� ��ġ
^F2::
    InputBox,num,���� �� ����,���� ���� �Է� �Ͻÿ� �⺻ ������ 2�� �Է��Ͻÿ�.
    InputBox, Npart, part column choice, part column �Է��Ͻÿ� (B or F ���)
    InputBox, Qty, Qty column choice, Qty column �Է��Ͻÿ� (B or F ���)
    InputBox, nnnn, Qty column choice, NNN column �Է��Ͻÿ� (B or F ���)

    Loop
    {	
        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
        while (tmp1)
        {
			click ,1113,198
			Sleep,300
			click,731,237
			Sleep,300
			Send {down} {down} {enter}
			
            VAR2=%Npart%%num% 
            var3=%Qty%%num% ;   ����
			var4=%nnnn%%num%
			part := Xl.Range(VAR2).value
			qt := round((Xl.range(var3).value),0)
            if (part="")
            {
                MsgBox,���α׷� ����
                ExitApp
            }
            Click,181,264
            send %part%
            Sleep,300
            send {enter}
            sleep 2000 
			
			Click,1188,648
			Send %qt%
			Click,1188,648,2
			
			pause
			Xl.Range(var4).value := "OK"
            num++		
            sleep 500
        }
    }
Return

^Space::Pause

^PGUP::Reload

^PGDN::ExitApp 