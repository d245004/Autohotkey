#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1

;SendMode Input
XLS_file_path3 := A_WorkingDir, "8월10일 기아 페기건 입력 할 것.xlsx"
Xl := ComObjCreate("Excel.Application")
Xl := ComObjActive("Excel.Application")

num=2
; XLS_file_path3 := A_WorkingDir . "HV11.XLS"
; X1 := ComObjCreate("Excel.Application")
; X1 := ComObjActive("Excel.Application")
; X1.Range("A:H").NUMBERFORMAT := "@"   ; ���� ��� TEXT�����

;HV11 대리점 불용재고 폐기품목 입력
^F2::
InputBox,num,시작 번호를 입력하시요
Loop 1000
{
	WinGetActiveTitle,tmp1
	while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
	{
			Click,1267,203
			Sleep,1000
			cur = 330
		Loop 15
		{
					VAR=C%num%            ; PART
					VAR2=E%num%           ; QTY
					var7 = K%num%
					part := Xl.Range(VAR).value
					if part =
							break

					qty := round(Xl.range(var2).value,0)
					Sleep 1000
					Click,180,%cur%
					sleep 300
					send %part%
					Sleep 300
					send {tab}
					send %qty%
					sleep 300
					send {tab}
					cur += 23
					num++
					Xl.Range(var7).value := "OK"
					if (A_Index >15)
						{
							break
						}
		}

		Click,1388,203

		if part =
		{
			MsgBox,데이터가 없네. 작업을 종료해야지
			ExitApp
		}
		Sleep,500
		Pause
		Send {Enter}
		Sleep,2000
		Send {Enter}
		Sleep,2500
	}
}




return

^Space::Pause
^PGDN::ExitApp