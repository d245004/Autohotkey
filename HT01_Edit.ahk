#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1

;FileSelectFile, path
;Xl := ComObjCreate("Excel.Application")
;Xl.Workbooks.Open(path)
;Xl.Visible := True

XLS_file_path3 := A_WorkingDir . "HT01.xls"
Xl := ComObjCreate("Excel.Application")
Xl := ComObjActive("Excel.Application")


;매출전표 수정 프로그램 (ABC  삭제)

Send {ctrl down}{f2}{ctrl up}

^F2::
WinActivate ahk_class Chrome_WidgetWin_1
send +^{NumpadAdd}

InputBox,num,,Start Number Input!
Loop
{
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1
	while (tmp1)
	{
        var = C%num%
        var2 = F%num%
        jun_no := Xl.Range(var).value
        if jun_no =
            {
                MsgBox, Work Out
				ExitApp
            }
		Click,929,203
		Sleep,1000
		Click,261,234
		Sleep,1000
        Send %jun_no%
        Sleep,1000
		Click,997,203
		Sleep,1000
		Click,561,237
        Sleep,1000
		Click,195,442,2
		Sleep,1000
        Send {DEL}
		Sleep,1000
        Click,1056,204
        Sleep,3000
		Xl.Range(var2).value := "OK"
		sleep,1000
		num++
    }
}
return

^Space::Pause

^PgUp::Reload

^PGDN::ExitApp