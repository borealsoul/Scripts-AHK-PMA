/************************************************************************
 * @description SCRIPTS QUE USO NO TRABALHO
 * @file SELECT.ahk
 * @author Aurora Júlia Moreira 
 * @since 09/03/2023
 * @date 2023/11/01
 * @version 0.1.5
 ***********************************************************************/

#SingleInstance

; Alt + Clique
; 	Selecionar múltiplas dívidas.
; Alt + Shift + Clique
; 	Selecionar múltiplas dívidas, até o fim da tela.
; Alt + Shift + C
; 	Seleção de parcelas para impressão.
; Ctrl + Shift + V
; 	Colar dentro do Terminal
; F8
; 	Digitação de Guias
; Shift + F8
;	Digitação de Guias, com impressão de carnê e termo.
; Shift + F5
; 	F5 já com pesquisa do nome.
; F6
;	Extração de CDA.
; F4
;	Cancelamento de parcelamento por inadimplência.

; Alt + 1 no Word
;	Exportar em PDF para a pasta da data de hoje.

#HotIf WinActive("Terminal remoto — Mozilla Firefox") or WinActive("Terminal remoto - Google Chrome")
; Se estiver no Terminal Remoto da Sonner

; Selecionar várias dívidas de uma vez.
!LButton::		; Alt + Clique
!+LButton:: {	; Alt + Shift + Clique
	shift_pressionado := False
	if GetKeyState("Shift")
		shift_pressionado := True

	; TODO: #1 Deveria ser possível calcular a distância necessária entre os blocos
	; pela resolução da tela. Até agora só testei no computador que uso (1366) e no da Camila (1920)
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

; Seleção de parcelas para impressão.
+!C:: {		; Alt + Shift + C
	; TODO #2 Não é preciso a diferenciação se tem entrada ou não.
	entrada_parcelamento := MsgBox("Este parcelamento possui uma entrada?", , "YesNoCancel Icon?")
		if entrada_parcelamento != "Cancel" {
			qnt_parcelas := 14 - A_MM
			if entrada_parcelamento == "No"
				qnt_parcelas--

			Send "+{TAB 2}{SPACE 2}{TAB}"
			Sleep 200

			loop qnt_parcelas
				Send "{DOWN}{SPACE}"
			Sleep 200

			; Caixa de mensagem para conferir se o parcelamento é para o próprio mês ou para o próximo.
			venc_parcelamento := MsgBox("O parcelamento é para o mês vigente? (" A_MMMM ")", , "YesNo Icon?")
			
			if venc_parcelamento == "No"
				Send "{SPACE}"

			Send "!A{ENTER}"
		}
}

; Colar dentro do terminal
^+V::Send A_Clipboard	; Ctrl + Shift + V

; Digitação de Guias
F8::		; F8
+F8:: {		; Shift + F8
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

; Roda do Mouse como Setas do Teclado
WHEELDOWN::Send "{DOWN}"
WHEELUP::Send "{UP}"

; F5 já com pesquisa por nome
+F5:: {		; Shift + F5
	Send "{F5}!{T}{F3}"
	Send "+{TAB 2}{RIGHT}{TAB 2}"
}

; Extração de CDA
F6:: {
	Sleep 150
	Send "!V"
	Sleep 20
	Send "{C}"
	Sleep 20
	Send "{DOWN 8}{ENTER}"
}

; Cancelamento de Parcelamento
; TODO: #5 Adicionar um 4o botão para cancelar em caso de simulação de parcelamento (ou torná-lo outro tipo de mensagem?)
F4:: {
	SetTimer ChangeButtonNames, 40
	cancelamento_autorizado := MsgBox("Gostaria de cancelar o parcelamento selecionado por qual motivo?", "Cancelamento de Parcelamento", "Y/N/C Icon!")
	if (cancelamento_autorizado != "Cancel") {
		Sleep 50
		Click "Right"	
		Sleep 50
		Click 10, 65, "RELATIVE" 
		Sleep 100
		Send "{ENTER}"
		Send "+{TAB}"
		if (cancelamento_autorizado = "Yes")
			Send "Inadimplencia."
		else
			Send "Para baixa de SISBAJUD."
		Send "!{I}"
	}	

	ChangeButtonNames() {
		if !WinExist("Cancelamento de Parcelamento")
			return
		SetTimer , 0
		WinActivate
		ControlSetText "&Inadimp.", "Button1"
		ControlSetText "&SISBAJUD", "Button2"
	}
}

; No Word.
#HotIf WinActive("ahk_exe WINWORD.EXE")
Alt::Return
; No Word, apertando Alt.
#HotIf GetKeyState("Alt") and WinActive("ahk_exe WINWORD.EXE")
!1:: {	; Alt + 1
	A_Clipboard := ""
	Send "^{c}"
	ClipWait
	A_Clipboard := StrUpper(A_Clipboard)
	Send "!{1}"
	Sleep 300
	; TODO #3 - É preciso tornar o caminho de pasta dinânico também.
	Send "{F4}^{A}{BACKSPACE}D:\PREFEITURA\Documents\ARQUIVOS\11 - Novembro\" A_YYYY "-" A_MM "-" A_DD "{ENTER}{TAB}"
	Sleep 200
	Send "!{N}" A_Clipboard "{ENTER}"
	Sleep 1800
	Send "!{F4}"
}



