SendMode Input             ;  send 명령을 빠르게 실행하기위해서 사용
;Coordmode, Mouse, Screen   ;  전체화면에서 마우스 포인트를 사용한다고 설정하는 것
num=2
XLS_file_path3 := A_WorkingDir . "MACHUL.xls"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:P").NUMBERFORMAT := "@"


; HAIMS 세금계산서  프로그램
F2:: 

InputBox,nal,발행 일자 입력,발행 일자를 입력 하세요. ex-> 20190331

Loop 1000
{	
	WinGetActiveTitle,tmp1
		while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
	{
		VAR2=C%num%   ;    C열의 몇 번째   CODE
		var3=M%num%   ;    세전금액
		var4=N%num%   ;    부가세 
		code := X1.Range(VAR2).value
		price := round(X1.Range(var3).value,0)
		vat := round(X1.Range(var4).value,0)
		bigo := price + vat
		if (code = "")
			{
				MsgBox,작업 완료! (지루했네요.)
				ExitApp
			}
		Click,969,203
		Sleep,300
		Click,406,266,2
		Sleep,300
		Send,{del}
		Sleep,300
		Send,%nal%
		Sleep,300
		Click,245,371,2
		Sleep,300
		Send,{del}
		Sleep,300
		Send,%code%
		sleep 1300
		Send {down}
		send {enter}
		sleep 1300     
		Click,828,627,2
		Sleep,300
		Send,{del}
		Sleep,300
		Send,%price%
		Sleep,300
		Send,{tab}
		Sleep,300
		;  Click,987,625
		Send,%price%
		sleep,300
		Send,{tab}
		Sleep,300
		send,%vat%
		Sleep,300
		Send,{enter}
		Sleep,300
		;MouseMove,1177,779
		;Sleep,300	
		Click,1177,779,2
		Sleep,300
		Send,{del}
		Sleep,300
		send,%bigo%
		Sleep,300
		Send,{enter}
		Sleep,300
		MouseMove,1100,202
		Loop
			{
				getkeystate,vvar,enter
				IF vvar=D
				{
					Click,1100,202
					Sleep,300
					Send,{enter}
					; Click,842,174
					Sleep,3000
					Send,{enter}
					; Click,917,173
					Sleep,1000
					break
				}
			}
			
		num++	
		Pause
}
Pause  ;잠시 중지
}

return    ; 프로그램 원위치로

PGUP::Pause   ; 누르면 잠시 중지를 ON / OFF 역활을 한다

PGDN::ExitApp   ; 프로그램을 종료한다, 메모리에서 삭제된다