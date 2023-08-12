#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

; CoordMode, Mouse,Caret, Screen,Pixel

; XLS_file_path3 := A_WorkingDir . "부품별 수불 집계(20211027).xls"
; Xl := ComObjCreate("Excel.Application")
; Xl := ComObjActive("Excel.Application")

FileSelectFile, path
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True
; Xl 과 Xl 구분 할 것 (에러 발생 이유)

^F2::
    InputBox,num,시작 행 선택,시작 행을 입력 하시요. 기본 시작은 3을 입력하시요.
    InputBox,s_lep,LEP 선택,LEP 행을 입력 하시요.
    InputBox,s_part,PART 행 선택,PART 행을 입력 하시요.

    aa = 1
    Loop
    {	
        tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
        while (tmp1)
        {
            VAR2=%s_part%%num% ;   3143029200 A열의 몇 번째   part number
            part := Xl.Range(VAR2).value
            if (part="")
            {
                MsgBox,Press any key
                return
            }
            var6=H%num% ;  E열  PART NAME
            var10=I%num% ;  D열  doprice
            var11=J%num%
            var17=%s_lep%%num% ;  C열  lep
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
            sleep 1500 ;   1000은 1초를 의미한다

            Clipboard=
            Click,235,600,2
            Sleep,500
            Send ^c
            clipwait,0
            Xl.Range(var10).value := Clipboard ; doprice D열

            Clipboard=
            MouseClickDrag,L,416,336,607,336 ;ENGLISH
            ;~ MouseClickDrag,L,710,336,910,336    ;KOREAN
            Sleep,500
            Send ^c
            clipwait,0
            Xl.range(var6).value := Clipboard ;   partname E열

            Clipboard=
            MouseClickDrag, L, 309, 909, 421, 909
            Sleep,500
            Send ^c
            clipwait,0
            Xl.Range(var11).value := Clipboard ; 로케이션


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

^Space::Pause ; 누르면 잠시 중지를 ON / OFF 역활을 한다

^PgUp::Reload

^PGDN::ExitApp ; 르로그램을 종료한다, 메모리에서 삭제된다
