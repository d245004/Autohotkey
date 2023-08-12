 SendMode Input   ;  send 명령을 빠르게 실행하기위해서 사용
CoordMode, Mouse,Caret, Screen
XLS_file_path3 := A_WorkingDir . "jaego.XLSx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:H").NUMBERFORMAT := "@"   ; 엑셀 셀을 TEXT로 고정함
return

^F2::
WinMove, 파트별재고 - Brave, , 0, 0, 494, 767	
InputBox,num,시작번호 입력,시작 번호를 입력하세요. (기본은 2로 입력)
if ErrorLevel
	ExitApp


Loop 
{	
	Loop
	{
		tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
		while (tmp1)
		{
			; 엑셀 자료 대입시 := 연산자는 쓰지 않는다
			VAR2 = A%num%           ; PART  지정
			var5 = B%num%           ; 수량 지정
			var1 = C%num%           ; OK 아니면 재고 부족
		
			part := X1.Range(VAR2).value
			qty := round(X1.range(var5).value,0)
			if part =    ;변수 내용이 null이면
				{
					MsgBox,계속 작업하려면 F2 key press.. 
					return
				}
			
			Sleep,500
			Click,412,682
			sleep,500
			Click,246,418
			Sleep,500
			Click,288,130
			Sleep, 500
			Click,246,418
			Sleep,500
			Click,107,126
			Sleep,500

			Send %part%
			Sleep,1000
			Send {enter}
			Sleep,1000
			Click,189,387
			Sleep,1500

			ImageSearch,ax,ay,20,250,470,400, *80 loc.png
			if(errorlevel = 1)
				{
					MsgBox, 이미지서치 실패! 계속 작업하려면 F2 key press
					return
				}

			MouseClick,L,ax+220,ay+40,2,0  ; LOC Click
			Sleep,1500       

			Click,300,690    ; 조정 Button Click
			Sleep,3000
			tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
			while (tmp1)
			{
				Send, %qty%
				Sleep,500
				Send {enter}
				Sleep,500
				Send {enter}
				Sleep,500
				Click,242,410
				Sleep,500
				X1.range(var1).value := "-= OK  =-"
        	 	Sleep,500
				SoundPlay,click.mp3,wait
				break
			}
			num++
		}
	}
}
return

^Space::Pause

^PGDN::ExitApp

^PgUp::Reload

