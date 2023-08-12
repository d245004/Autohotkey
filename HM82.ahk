num=2
XLS_file_path3 := A_WorkingDir . "TEST.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:C").NUMBERFORMAT := "@"
return

;대리점 선순환지원 품목 요청 및 확정

^Home::
Loop 1000
{	
Loop 14
{
	WinGetActiveTitle,tmp1
		while (tmp1 = "HYUNDAI MOBIS Agent Inventory Management System - Chrome")
	{
		
		VAR=E%num%
		VAR2=I%num%
		part := X1.Range(VAR).value
		qty := round(X1.range(var2).value,0)
		send %part%
		sleep 1000
		send {tab}
		send %qty%
		sleep 100
		send {tab}
		;  Send {TAB}
		;  Send {TAB}
		num++
		;Pause
	    break
	}
}
Pause
}

return

^F2::Pause

^end::ExitApp