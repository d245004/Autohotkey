 SendMode Input   ;  send ����� ������ �����ϱ����ؼ� ���
CoordMode, Mouse,Caret, Screen
XLS_file_path3 := A_WorkingDir . "jaego.XLSx"
X1 := ComObjCreate("Excel.Application")
X1 := ComObjActive("Excel.Application")
X1.Range("A:H").NUMBERFORMAT := "@"   ; ���� ���� TEXT�� ������
return

^F2::
WinMove, ��Ʈ����� - Brave, , 0, 0, 494, 767	
InputBox,num,���۹�ȣ �Է�,���� ��ȣ�� �Է��ϼ���. (�⺻�� 2�� �Է�)
if ErrorLevel
	ExitApp


Loop 
{	
	Loop
	{
		tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
		while (tmp1)
		{
			; ���� �ڷ� ���Խ� := �����ڴ� ���� �ʴ´�
			VAR2 = A%num%           ; PART  ����
			var5 = B%num%           ; ���� ����
			var1 = C%num%           ; OK �ƴϸ� ��� ����
		
			part := X1.Range(VAR2).value
			qty := round(X1.range(var5).value,0)
			if part =    ;���� ������ null�̸�
				{
					MsgBox,��� �۾��Ϸ��� F2 key press.. 
					return
				}
			
			Sleep,500
			Click,412,682
			sleep,500
			Click,246,418
			Sleep,500
			Click,288,130
			Sleep, 500
			Click,246,418
			Sleep,500
			Click,107,126
			Sleep,500

			Send %part%
			Sleep,1000
			Send {enter}
			Sleep,1000
			Click,189,387
			Sleep,1500

			ImageSearch,ax,ay,20,250,470,400, *80 loc.png
			if(errorlevel = 1)
				{
					MsgBox, �̹�����ġ ����! ��� �۾��Ϸ��� F2 key press
					return
				}

			MouseClick,L,ax+220,ay+40,2,0  ; LOC Click
			Sleep,1500       

			Click,300,690    ; ���� Button Click
			Sleep,3000
			tmp1 = WinWaitActive,ahk_class Chrome_WidgetWin_1 
			while (tmp1)
			{
				Send, %qty%
				Sleep,500
				Send {enter}
				Sleep,500
				Send {enter}
				Sleep,500
				Click,242,410
				Sleep,500
				X1.range(var1).value := "-= OK  =-"
        	 	Sleep,500
				SoundPlay,click.mp3,wait
				break
			}
			num++
		}
	}
}
return

^Space::Pause

^PGDN::ExitApp

^PgUp::Reload

