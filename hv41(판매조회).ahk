 #SingleInstance, Force
 SendMode Input
SetWorkingDir, %A_ScriptDir%
;Coordmode, Mouse, Screen   ;  전체화면에서 마우스 포인트를 사용한다고 설정하는 것
XLS_file_path3 := A_WorkingDir . "CRP 대리점별 반품 요청 --기아.xls"
X1 := ComObjActive("Excel.Application")
X1.Range("A:K").NUMBERFORMAT := "@"

^F2::
InputBox,num,시작 행 선택,시작하려면 기본 시작 2를 입력하시요


aa = 1
Loop
{	
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
	while (tmp1)
	{
		VAR2=B%num%   ;   A열의 몇 번째   part number
		part := X1.Range(VAR2).value
		if (part="")
		{
			MsgBox,Press any key
			return
		}
		var4=L%num%     ;매입수량
		maip := X1.Range(var4).value

        var11=Q%num%

	    Click,193,263,2
		Sleep,500
		Send, {Home}
		Sleep, 500
		Send {Delete 20}
		Sleep,500
		send %part%
		Sleep,500
		Send, {Enter}
		sleep 1500     ;   1000은 1초를 의미

		Clipboard=
		Click,144,424,2
		Send, ^c
		ClipWait, 0
		jqty := Clipboard   ; 항목수
		sleep,300



		Clipboard=
		Click,273,423,2
		Send, ^c
		ClipWait, 0
		total := Clipboard    ; 판매수량
		Sleep, ,300
		sale := 

		ms := 0

		aa := 1
		While aa<=(jqty)
		{
			Clipboard=
			MouseClickDrag,L,508,495+ms,656,495+ms
			Sleep,300
			Send ^c
			clipwait,0
			sale := sale Clipboard ","
			Sleep, 300

			; Clipboard=
			; MouseClick, L, 770, 495+ms,2
			; Sleep, 300
			; Send, ^c
			; ClipWait, 0
			; sale := sale Clipboard ","
			; Sleep, 300

			aa += 1
			if (aa>13)
				Break
			ms += 22
		}

		X1.range(var4).value := "판매" total "개"
		X1.range(var11).value := sale
        

		Sleep,300
		num++		
		sleep 500
		aa += 1
		if (aa > 1300)
			{
				aa = 1
				Pause
			}
		; Pause
		Sleep, 1000
	}
}


^Space::   Pause

PGUP::     Reload

PGDN::ExitApp   ; 르로그램을 종료한다, 메모리에서 삭제된다
