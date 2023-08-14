#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1


; 엑셀 파일 열어놓고 작업하기

; XLS_file_path3 := A_WorkingDir . "TEST.xlsx"
; Xl := ComObjCreate("Excel.Application")
; Xl := ComObjActive("Excel.Application")

; 엑셀파일 선택창 열기 예제

; FileSelectFile, path
; FileSelectFile, path ,,D:/aaa_DownLoad/,,*.xlsx
; if path = 
; 	ExitApp
; Xl := ComObjCreate("Excel.Application")
; Xl.Workbooks.Open(path)
; Xl.Visible := True

; test.xlsx 만 지정해서 열기
path := "D:/aaa_DownLoad/test.xlsx"
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Open(path)
Xl.Visible := True




Gui Add, Text, x32 y10 w120 h23 +0x200, 거래처 코드
Gui Add, Text, x32 y56 w120 h23 +0x200, 시작 행 번호
Gui Add, Text, x32 y106 w120 h21 +0x200, 입력 품목 건수
Gui Add, Edit, x163 y11 w120 h21 vSANG, Z00
Gui Add, Edit, x163 y57 w120 h21 vnum, 2
Gui Add, Edit, x163 y105 w120 h21 vsu, 60
Gui Add, Button, x80 y145 w161 h23 gBtn, 작업 시작

Gui Show, x10 y10 w310 h181, Window
Return

GuiEscape:
GuiClose:
	Xl.activeWorkbook.save()
	Xl.Quit
    ExitApp

Btn:
	Gui, Submit, NoHide
	WinActivate ahk_class Chrome_WidgetWin_1
; 윈도우 고정하기
    send +^{NumpadAdd}

	;HT01 매출 전표 입력
	;Click,1363,205
	Loop 1000
	{
		WorkWindow = WinWaitActive,ahk_class Chrome_WidgetWin_1
		while (WorkWindow)
	{
			Click,926,202
			Sleep,1000
			Click,453,358
			Sleep,1000
			Send %SANG%
			Sleep,2000
			Send {DOWN}
			Send {ENTER}
			Sleep,2000
			Click,195,442 ;내용선택
			;Sleep,1000
        	;Send ABC
			Sleep,2000
			Send {ENTER}
			Sleep,1000
			Click,127,506,2
			Send {DEL}


		Ipsu = 0
		Loop %su%
		{
			VAR=B%num%            	; LEP
			VAR2=C%num%           	; PART
			var5=E%num%           	; SU
			var6=I%num%				; count
			var7=J%num%				; "OK"
			part := Xl.Range(VAR2).value
			if part =    ; Work Close
				break
			qty := round(Xl.range(var5).value,0)
			lep := Xl.Range(VAR).value
			send %lep%
			Sleep 300
			send {Tab} %part% {Tab}
			sleep 1000
			send %qty%
			sleep 300
			send {Tab} {Tab} {Tab}
			Ipsu++
			cap = %Ipsu%
			Xl.Range(var6).value := Cap
			Xl.Range(var7).value := "OK"
			num++
;			if (A_Index > 60)  ; loop문의 순환 횟수 카운트
;				break
		}
		send {enter}
		Quit()
	}

	Quit()
	{
		MsgBox,작업을 종료 합니다.
		;Pause
		;Reload
		Exit
		;return
	}
}


^Space::Pause

^PgUp::
	Xl.Application.DisplayAlerts := False
	Xl.ActiveWorkbook.Close
	Xl.Visible := False
	Xl.Quit
	Reload

^PGDN::
	Xl.ActiveWorkbook.save()
	Xl.Quit
	ExitApp

