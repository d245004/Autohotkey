SendMode Input   ;  send 명령을 빠르게 실행하기위해서 사용
CoordMode, Mouse, Screen


^F12::   ;  F12
	Loop
	{
;		WinGetActiveTitle,tmp1
		;while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
		
;		while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System")
;		- 개인 - Microsoft Edge")
		tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
		while (tmp1)
		{
			Click,1647,201  ;저장버튼 클릭
			Sleep,1000
			Click,1280,173
			Sleep,3000
			Click,1353,171
			Sleep,1000
			MouseMove,1524,202
			Loop
			{
				getkeystate,VAR,enter
				IF VAR=D
				{
				Click,1524,202
				break
				}
			}
			return
		}
	}
	
  
^Del::   ;  ctrl + shift + Del
Loop
	{
;		WinGetActiveTitle,tmp1
;		while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
		tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
		while (tmp1)
		
		
		{
			MouseGetPos, xpos, ypos    ; 마우스 위치 기억
			old = ypos-24
			;MsgBox,%ypos%
			Click, %xpos%, %ypos%,2
			Sleep,500
			;Send ^c
			;Sleep,500
			;Send ^v
			;Sleep,500
			Click,1756,304
			Sleep,1500
			Click,1285,174
			Sleep,1500
			;MouseClick,Left, %xpos%, %ypos% ,2  ; 기억한 위치로 이동
			Click,Left, %xpos%, %ypos% ,2  ; 기억한 위치로 이동
			;MouseMove,1885,301

			return
		}
	}
	
^+PGUP::Pause    ;  ctrl + shift + PGUP
	
^+PGDN::ExitApp  ; ctrl + shift + PGDN
 	

^+F9::           ; ctrl + shift + F9
	{
		;WinGetActiveTitle,tmp1
		;while (tmp1= "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
		tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
		while (tmp1)
		{
			aa:=Array()
			aa[1]:="HP21"
			aa[2]:="HT01"
			aa[3]:="HT11"
			count=1
			
			Send, ^+z
			Sleep,500
			Send,^1  ; 1번창 이동
			Sleep,500
			Click,510,1020 ; 모두 지우기
			Sleep,500
			
			loop,3
			{
			Click,469,134
			Send,{DEL 5}
			Sleep,500
			Send,% aa[count]
			Sleep,500
			Send,{ENTER}
			if (aa[A_Index]="HT01") ; 배열 변수 복잡하기는 하네
				Sleep,10000
			else
				Sleep,3000
			count++
			}
			Send,^2   ; 2번창 이동

			aa:=Array()
			aa[1]:="HI01"
			aa[2]:="HI02"
			aa[3]:="HR01"
			aa[4]:="HR07"
			aa[5]:="HR42"
			aa[6]:="HV41"
			aa[7]:="HV42"
			aa[8]:="HR44"
			aa[9]:="HR45"
			count=1

			Sleep,500
			Click,510,1020   ; 모두 지우기
			sleep,500
			
			loop,9
			{
			Click,469,134
			Send,{DEL 5}
			Sleep,500
			Send,% aa[count]
			Sleep,500
			Send,{ENTER}
			If (aa[A_Index]="HI01")   ; 배열 변수 복잡하기는 하네
				Sleep,10000
			else
				Sleep,3000
			count++
			}
			
			Send,^1  ; 1번 창으로 이동
			Click,706,1010
			Sleep,300
			Click,522,303
			
			return
		}
	}
	
+SC039::Send {vk15sc138}   ; L쉬프트 + 스페이스    (한영 변환)
^SC039::Send {vk19sc11D}   ; L Ctri + Space       (한자 변환)   무척 편안하군. 공부해야되

::TPRMA::
	Send,계산서 발행, 부가세  원 받아야한다. 
	Send,{BACKSPACE}
	Send,{LEFT}
	Send,{LEFT}
	Send,{LEFT}
	Send,{LEFT}
	Send,{LEFT}ccccccc
	Send,{LEFT}
	Send,{LEFT}
	return
	
