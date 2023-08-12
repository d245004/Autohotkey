#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1

Gui Add, Button, x16 y8 w253 h38 gBtn1, HR01 견적서 조회
Gui Add, Button, x16 y56 w253 h38 gBtn2, HR03_new 청구 수량 조정
Gui Add, Button, x16 y104 w253 h38 gBtn3, HR07(주문2) 전문점 수령으로 주문
Gui Add, Button, x16 y152 w253 h38 gBtn4, HR07 지원센타 청구 입력
Gui Add, Button, x16 y200 w253 h38 gBtn5, HT01 매출 전표 입력
Gui Add, Button, x16 y248 w253 h38 gBtn6, HT01-Edit 매출 전표 수정 입력 (ABC)
Gui Add, Button, x16 y296 w253 h38 gBtn7, HV41 이상 수요 조회
Gui Add, Button, x16 y344 w253 h38 gBtn8, jaego_PDA  PDA 재고 수정
Gui Add, Button, x16 y392 w253 h38 gBtn9, KIA_J 부품 도가, 부품명, 로케이션 조회
Gui Add, Button, x16 y440 w253 h38 gBtn10, PDA302 PDA 재고 빼기
Gui Add, Button, x16 y488 w253 h38 gBtn11, PDA 재고 조정
Gui Add, Button, x16 y536 w253 h38 gBtn12, Plus PDA PDA 재고 더하기
Gui Add, Button, x16 y584 w253 h38 gBtn13, VAT 조회   보험건 VAT 양식 만들기
Gui Add, Button, x16 y632 w253 h38 gBtn14, 재물조사   PDA 재물조사

Gui Show,x10 y10 w286 h687, AutoHotKey
Return

GuiEscape:
GuiClose:
    ExitApp


Btn1:
Gui, Hide
RunWait,hr01.ahk
Gui, Show
return

Btn2:
Gui, Hide
RunWait,hr03_new.ahk
Gui, Show
return

Btn3:
Gui, Hide
RunWait,hr07(주문2).ahk
Gui, Show
return

Btn4:
Gui, Hide
RunWait,hr07.ahk
Gui, Show
return

Btn5:
Gui, Hide
RunWait,ht01.ahk
Gui, Show
return

Btn6:
Gui, Hide
RunWait,ht01_edit.ahk
Gui, Show
return


Btn7:
Gui, Hide
RunWait,hv41.ahk
Gui, Show
return

Btn8:
Gui, Hide
RunWait,jaego_pda.ahk
Gui, Show
return


Btn9:
Gui, Hide
RunWait,kia_j.ahk
Gui, Show
return

Btn10:
Gui, Hide
RunWait,pda302.ahk
Gui, Show
return


Btn11:
Gui, Hide
RunWait,pda.ahk
Gui, Show
return

Btn12:
Gui, Hide
RunWait,plus pda.ahk
Gui, Show
return


Btn13:
Gui, Hide
RunWait,vat 조회.ahk
Gui, Show
return


Btn14:
Gui, Hide
RunWait,재물조사.ahk
Gui, Show
return


^!PGUP::
Reload

^!PGDN::
ExitApp
