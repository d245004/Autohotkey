CoordMode, Mouse, Screen

num=82
XLS_file_path3 := A_WorkingDir . "test.XLSX"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
;X1.Range("A:D").NUMBERFORMAT := "@"   ; 엑셀 셀을 TEXT로 고정함
return


;HV01  조정 입력 관리
^F2::
Loop 1000
{	
		Click,570,335,2
		Sleep 1000
		Send {DELETE}
Loop 20
{
;	WinGetActiveTitle,tmp1
;		while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
		tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
		while (tmp1)

{
		VAR=B%num%            ; A열 지정   LEP
		VAR2=C%num%           ; B열 지정   PART NUMBER
		var4=J%num%           ; C열 지정   QTY
		var5=K%num%           ; D열 지정   "입력 완료"
		part := X1.Range(VAR2).value
;		lep := substr(part,1,1)  ;변수에서 지정한부분 텍스트만 뽑아냄
;		part := substr(part,2,15)  ;;변수에서 지정한부분 텍스트만 뽑아냄
		;MsgBox,%part%
		if part =             ; PART 내용이 null이면
			{
				MsgBox,종료합니다
				ExitApp
			}
		qty := round(X1.range(var4).value,0)
		lep := X1.Range(VAR).value
		send %lep%		;  LEP 선택
		send {tab}      ;  PART로 이동
		send %part%     ;  PART 입력
		sleep 2000
		send {tab}      ;  창고로 이동
		sleep 1400
		send {tab}		;  조정 수량으로 이동
		send %qty%		;  조장 수량 입력
		Sleep 1500
;		sleep 3500       ;  2.5초 대기   (줄이면 ERROR 발생)
		send {tab}		;  조정단가로 이동
		sleep 1500      ;  1초 대기
		Send {tab}      ;  다음 줄 LEP로 이동
		X1.Range(var5).value := "입력 완료"
		
		num++
	    break
	}
}
Click,1797,210
;Sleep 20000
Pause
Click,1672,210
Sleep 500
}

return
^Space::Pause

^PGUP::Reload

^PGDN::ExitApp