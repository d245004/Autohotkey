CoordMode, Mouse,Caret, Screen
XLS_file_path3 := A_WorkingDir . "list_song.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:H").NUMBERFORMAT := "@"   ; ???? ???? TEXT?? ??????
return

;???? ???????? ??? ??????

F2::
InputBox,num,???? ?? ????,???? ???? ??? ????. ?? ?????? 2?? ???????.
Loop
{	
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
	while (tmp1)
	{
		var1=A%num%   ; code
		var2=B%num%   ;  ??????
		var3=D%num%  ;  ???????
		var4=E%num%   ;  ??????
		var5=F%num%   ;  ???

		part := X1.Range(var1).value
		sang := X1.Range(var2).value
		a18num := X1.Range(var3).value
		name := X1.Range(var4).value
		addre := X1.Range(var5).value
		if (part="")
		{
			MsgBox,???��?? ????
			ExitApp
		}
		Click,161,233,2
		Sleep,500
		Send {del}
		Sleep,1500
		Send %part%
		Sleep,500
		
		MouseClickDrag,L,350,233,549,233
		Sleep,500
		Send {del}
		Sleep,500
		Send %sang%
		Sleep,500
		
		Click,217,299,2
		Sleep,500
		Send {del}
		Sleep,500
		Send %a18num%
		Sleep,500
		
		MouseClickDrag,L,208,330,373,330
		Sleep,500
		Send {del}
		Sleep,500
		Send %name%
		Sleep,500
		
		Click,647,449
		Sleep,500
		Click,300,447
		Send %addre%
		Sleep,500
		
		;Pause
		
		Click,1081,199
		Sleep,1500
		Click,848,188
		Sleep,1500
		
		num++
		Pause
		
	}
Pause  ;??? ????
}

return    ; ???��?? ???????

PGUP::Pause   ; ?????? ??? ?????? ON / OFF ????? ???

PGDN::ExitApp   ; ???��???? ???????, ?????? ???????