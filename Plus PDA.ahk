SendMode Input   

XLS_file_path3 := A_WorkingDir . "KIA_J.XLSX"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:H").NUMBERFORMAT := "@"   
return

F2::
;Send,^+{Q}

InputBox,num,시작번호 입력,시작 번호를 입력하세요. (기본은 2로 입력)
if ErrorLevel
	ExitApp
	Click,412,682   
	Sleep,700
	Click,246,418
	Sleep,3000

Loop 
{	
Loop 
{
	WinGetActiveTitle,tmp1
		while (tmp1 = "파트별재고 - Chrome")
	{
		VAR = C%num%            
		VAR2 = A%num%           
		var5 = B%num%           
		var1 = D%num%           
		
		part := X1.Range(VAR2).value
		qty := round(X1.range(var5).value,0)
		lep := X1.Range(VAR).value

		if part =    
			{
				MsgBox, 131120, 작업 중지 여부, 자료 없음. 종료합니다
				ExitApp
			}
		Sleep,3000	
		Click,277,127   
		Sleep,300
		

		Click,107,126
		Sleep,300
		Send,{DEL 15}
		Sleep,500
		Send %part%
		Sleep,700
		Send {enter}
		Sleep,700
		
		
		if (lep = "H")   
			Click,189,387
		else
			Click,296,411
		
		
		
		Sleep,700
		Click,237,340    
		Sleep,700       
		Click,300,690    
		Sleep,700
		WinGetActiveTitle,tmp2
			while (tmp2 = "재고조정 - Chrome")
			{
				Click,250,256,2
				Sleep,700
				Send {TAB}
				Send %qty%
				Sleep,700
				Send {enter}
				Sleep,700
				Send {enter}
				Sleep,700
				Click,247,418
				Sleep,700
				break
			}
		
		Sleep,700
		;Click,424,685
		;Sleep,500
		X1.range(var1).value := "PLUS OK"

		num++
		break
	}
}
}

return

PGUP::Pause    ; SPACE 키는 쓰면 안된다
	
		

PGDN::ExitApp