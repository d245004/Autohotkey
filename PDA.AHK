XLS_file_path3 := A_WorkingDir . "IMSI.XLSx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
; X1.Range("A:H").NUMBERFORMAT := "@"   ; 엑셀 셀을 TEXT로 고정함
; return

^F2::
WinMove, 파트별재고 - Brave, ,,, 494, 767	
InputBox,num,시작번호 입력,시작 번호를 입력하세요. (기본은 2로 입력)
if ErrorLevel
	ExitApp


Loop 
{	
	aa = 1
	Loop
	{
		tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
		while (tmp1)
		{
			; 엑셀 자료 대입시 := 연산자는 쓰지 않는다
			VAR = A%num%            ; LEP 지정
			VAR2 = B%num%           ; PART  지정
			var5 = C%num%           ; 수량 지정
			var1 = D%num%           ; OK 아니면 재고 부족
		
			part := X1.Range(VAR2).value
			qty := round(X1.range(var5).value,0)
			lep := X1.Range(VAR).value
			if part =    ;변수 내용이 null이면
				{
					MsgBox,계속 작업하려면 F2 key press.. 
					return
				}
			
			Sleep,1500
			Click,412,682
			sleep,700
			Click,246,418
			Sleep,700
			Click,288,130
			Sleep, 500
			; Click,246,418
			; Sleep,500
			Click,107,126
			Sleep,500

			Send %part%
			Sleep,1000
			Send {enter}
			Sleep,500
			if (lep = "H")   ; IF 문은 필히 ()로 싸매야 된다 (무려 반나절 걸림)
				Click,189,387
			else
				Click,296,411

			Sleep,500
			; ImageSearch,ax,ay,0,0,480,760, *80 loc.png
			; if(errorlevel = 1)
			; 	{
			; 		MsgBox, 이미지서치 실패! 계속 작업하려면 F2 key press
			; 		return
			; 	}

			; MouseClick,L,ax+270,ay+40,2,0  ; LOC Click
			click ,457,344,2
			Sleep,500       
			Send ^c
			Sleep,500
			clipwait,0
			if (qty > Clipboard)
			{
				X1.range(var1).value := "재고부족"
				Click,288,130
				Sleep,1000
				Click,412,682
				Sleep, 500
				num++
				Continue
			}

			Click,300,690    ; 조정 Button Click
			Sleep,3000
			tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
			while (tmp1)
			{
				Send, -%qty%
				Sleep,500
				Send {enter}
				Sleep,500
				Send {enter}
				Sleep,1500
				Click,242,410
				Sleep,1500
				X1.range(var1).Font.Bold:=1
				X1.range(var1).Font.Color:=255
				X1.range(var1).value := "-= OK  =-"
        	 	Sleep,1500
				SoundPlay,click.mp3,wait
				break
			}
			num++

			aa += 1
			if (aa>35)
				{
					MsgBox,계속 작업하려면 F2 key press.. 
					return
				}
				
			Pause
		}
	}
}
return

^Space::Pause

^PGDN::ExitApp

^PgUp::Reload

