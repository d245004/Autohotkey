CoordMode, Mouse,Caret, Screen


; XLS_file_path3 := A_WorkingDir . "1013보문상사-오토커머스 납기10월22일.xlsx"
; Xl := ComObjCreate("Excel.Application")
; Xl := ComObjActive("Excel.Application")
; Xl.Range("A:H").NUMBERFORMAT := "@"   ; 엑셀 셀을 TEXT로 고정함

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True


; 부품별 청구입력 및 조치
^F2::
    InputBox,num,시작 행 선택,시작 행을 입력 하시요 기본 시작은 4를 입력하시요.
    InputBox,code,CODE 선택,대리점 코드를 입력 하시요 (2450 or A041)
    InputBox, Npart, part column choice, part column 입력하시요 (B or F 등등)

    Loop
    {	
        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
        while (tmp1)
        {
            VAR2=%Npart%%num% ;    B열의 몇 번째   part number
            part := Xl.Range(VAR2).value
            if (part="")
            {
                MsgBox,프로그램 종료
                ExitApp
            }
            var6=%Npart%%num% ;  PART 
            ; var10=D%num%  ;  D열  doprice
            var3=F%num% ;  F열  모비스 현대재고 수량
            var4=H%num% ;  H열  오토넷 현대재고
            var7=G%num% ;  G열  모비스 기아재고
            var8=I%num% ;  I열  오토넷 기아재고
            var5=J%num% ;  J열  현대전문점 여부
            var9=K%num% ;  K열  기아전문점 여부 
            Click,181,264
            send %part%
            Sleep,300
            send {enter}
            sleep 3000 ;   1000은 1초를 의미한다

            Clipboard=
            Click,841,395,2
            Sleep,500
            Send ^c
            clipwait,0
            if (code = 2450)
                Xl.Range(var3).value := Clipboard ; 모비스현대재고 수량 F열
            else
                Xl.Range(var7).value := Clipboard ; 모비스기아재고 수량 G열

            Clipboard=
            Click,1210,451,2
            Sleep,500
            send ^c
            clipwait,0	
            if (code = 2450)
                Xl.Range(var4).value := clipboard ;  오토넷현대재고  H열
            else 
                Xl.Range(var8).value := clipboard ;  오토넷기아재고  I열

            Clipboard=
            ;MouseClickDrag,L,538,295,640,295
            Click,1203,294,2
            Sleep,500
            Send ^c
            clipwait,0
            if (code = 2450)
                Xl.range(var5).value := Clipboard ;   현대전문점 여부 J열
            else
                Xl.range(var9).value := Clipboard ;   기아전문점 여부 K열

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

^PGDN::ExitApp ; 르로그램을 종료한다, 메모리에서 삭제된다