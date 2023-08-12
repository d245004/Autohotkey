CoordMode, Mouse,Caret, Screen

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True


; 부품별 청구입력 및 조치
^F2::
    InputBox,num,시작 행 선택,시작 행을 입력 하시요 기본 시작은 2를 입력하시요.
    InputBox, Npart, part column choice, part column 입력하시요 (B or F 등등)
    InputBox, Qty, Qty column choice, Qty column 입력하시요 (B or F 등등)
    InputBox, nnnn, Qty column choice, NNN column 입력하시요 (B or F 등등)

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
            var3=%Qty%%num% ;   수량
			var4=%nnnn%%num%
			part := Xl.Range(VAR2).value
			qt := round((Xl.range(var3).value),0)
            if (part="")
            {
                MsgBox,프로그램 종료
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