CoordMode, Mouse,Caret, Screen
XLS_file_path3 := A_WorkingDir . "list_song.xlsx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:H").NUMBERFORMAT := "@"   ; ���� ���� TEXT�� ������
return


;��ü�� ��ǰ ���� ������ ����
F2::
InputBox,num,���� �� ����,���� ���� �Է� �Ͻÿ�. �⺻ ������ 2�� �Է��Ͻÿ�.
Loop
{	
	tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
	while (tmp1)
	{
		var1=A%num%   ; code
		part := X1.Range(var1).value
		if (part="")
		{
			MsgBox,���α׷� ����
			ExitApp
		}

		Click,1260,200

		Click,100,320
		Sleep, 1200
		Send, %part%

		Click,100,340
		Sleep, 1200
		Send, %part%

		Click,100,365
		Sleep, 1200
		Send, %part%

		Click,100,390
		Sleep, 1200
		Send, %part%

		Click,405,320
		Sleep, 1200
		Send, II
		
		Click,405,340
		Sleep, 1200
		Send, IN
		
		Click,405,365
		Sleep, 1200
		Send, IP
		
		Click,405,390
		Sleep, 1200
		Send, TC
		
		
		Click,715,326
		Sleep, 1200
		Send, -5

		Click,715,340
		Sleep, 1200
		Send, -5

		Click,715,365
		Sleep, 1200
		Send, -5

		Click,715,390
		Sleep, 1200
		Send, -5


		Click,1397,198
		Sleep, 2000

		Click,846,172
		Sleep, 2000

		; Click,919,168
		; Sleep, 2000

		num++
		Pause
		
	}
Pause  ;��� ����
}

return    ; ���α׷� ����ġ��

PGUP::Pause   ; ������ ��� ������ ON / OFF ��Ȱ�� �Ѵ�

PGDN::ExitApp   ; ���α׷��� �����Ѵ�, �޸𸮿��� �����ȴ�