#SingleInstance

#HotIf WinActive("Terminal remoto — Mozilla Firefox") or WinActive("Terminal remoto - Google Chrome")
!LButton::
!+LButton:: {
	shift_pressionado := False
	if GetKeyState("Shift")
		shift_pressionado := True

	switch A_ScreenHeight {
		case 768: distancia_mouse := 15.36
		case 1080: distancia_mouse := 18.3
	}

	vezes_clique := 9

	if shift_pressionado
		vezes_clique := 20

	Click
	loop vezes_clique { 
		if Mod(A_Index, 2) == 0
			MouseMove 0, (distancia_mouse * 1.3), 100, "RELATIVE"
		else
			MouseMove 0, (distancia_mouse * 0.7), 100, "RELATIVE"
		Click 
	}
	
	Sleep 50
	
	if shift_pressionado
		Send "{PgDn}"
}

+!C:: {	; -----------------------------SELEÇÃO DE CARNÊS PARA IMPRESSAO----------------------------Alt - Shift - C
	entrada_parcelamento := MsgBox("Este parcelamento possui uma entrada?", , "YesNoCancel Icon?")
		if entrada_parcelamento != "Cancel" {
			switch entrada_parcelamento {
				case "Yes":	qnt_parcelas := 14 - A_MM			; qnt_parcelas é igual a 14 - o mês atual
				case "No":	qnt_parcelas := 13 - A_MM
			}
			Send "+{TAB 2}{SPACE 2}{TAB}"		; Aperta Shift .. TAB x2 .. Espaço x2 .. TAB
			Sleep 200
			loop qnt_parcelas					; Loop pela qnt_parcelas
				Send "{DOWN}{SPACE}"			; Aperte seta para baixo .. Espaço 
			Sleep 200
			venc_parcelamento := MsgBox("O parcelamento é para o mês vigente? (" A_MMMM ")", , "YesNo Icon?")
												; Caixa de mensagem para conferir se o parcelamento é para o próprio mês ou para o próximo.
			switch venc_parcelamento {
				case "No": Send "{SPACE}"
			}
			Send "!A{ENTER}"					; Salva e fecha.
		}
}	; ---------------------------------------------------------------------------------------------

^+V::Send A_Clipboard	; -------------------------COLAR DENTRO DO SISTEMA-------------------------

F8::
+F8:: {
	shift_pressionado := False
	if GetKeyState("Shift")
		shift_pressionado := True

	Sleep 100
	Send "!V"
	Send "{D}"
	Sleep 20
	Send "{DOWN 10}{ENTER 2}"
	Sleep 20
	Send "+{TAB 8}{SPACE}"
	if shift_pressionado
		Send "+{TAB}{SPACE}"
		Send "{TAB 2}{SPACE}"
		Send "+{TAB}"
	Send "{TAB 8}"
}

WHEELDOWN::Send "{DOWN}"	; ---------------RODINHA DO MOUSE PARA SETAS DO TECLADO----------------
WHEELUP::Send "{UP}"		; ---------------------------------------------------------------------

+F5:: { 			; ---------------------F5 JÁ COM PESQUISA POR NOME--------------------Shift-F5
	Send "{F5}!{T}{F3}"
	Send "+{TAB 2}{RIGHT}{TAB 2}"
}

/* F4:: {
	Send "!{D}{RIGHT}"
} */

F6::	{	; ----------------------------EXTRAÇÃO DE CDA-----------------------------Ctrl - Shift - C
	Sleep 150
	Send "!V"
	Sleep 20
	Send "{C}"
	Sleep 20
	Send "{DOWN 8}{ENTER}"
}

F4:: {		; ---------------------------CANCELAMENTO DE PARCELAMENTO---------------------------------Ctrl - Alt - C
cancelamento_autorizado := MsgBox("Esse parcelamento será cancelado por inadimplência. Gostaria de prosseguir?", , "YesNo Icon!")
	if (cancelamento_autorizado = "Yes") {
		Sleep 100
		Click "Right"	
		Sleep 100
		Send "{DOWN 4}"
		Sleep 500
		Send "{ENTER}"
		Send "+{TAB}"
		Send "Inadimplencia."
		Send "!{I}"
	}	
}


/* #HotIf WinActive("ahk_exe WINWORD.EXE")
Alt::Return */

#HotIf GetKeyState("Alt") and WinActive("ahk_exe WINWORD.EXE")
!1:: {
	A_Clipboard := ""
	Send "^{c}"
	ClipWait
	A_Clipboard := StrUpper(A_Clipboard)
	Send "!{1}"
	Sleep 300
	Send "{F4}^{A}{BACKSPACE}D:\PREFEITURA\Documents\ARQUIVOS\10 - Outubro\" A_YYYY "-" A_MM "-" A_DD "{ENTER}{TAB}"
	Sleep 200
	Send "!{N}" A_Clipboard "{ENTER}"
	Sleep 1800
	Send "!{F4}"
}



