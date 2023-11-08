/************************************************************************
 * @description Passar as telas de dig. de guias, uma por uma
 * @file LISTADIGGUIAS.ahk
 * @author Aurora Júlia Moreira 
 * @since 08/11/2023
 * @date 208/11/202
 * @version 0.0.5
 ***********************************************************************/

#SingleInstance

; Recomendo usá-lo em tela dividida entre o navegador e o Excel, com e relatório em CSV.

F9:: {
	
    ; Array dos códigos das guias.
    lista := []
	
    ; Para cada índice e valor na lista
	for indice, valor in lista {
		Send "{ESC 2}{DEL 4}"   ; Apaga o número anterior
		Send valor              ; Envia o cód. da lista
		Send "{Enter 2}"        ; Entra na guia
		
		Send "!{TAB}"           ; Dá Alt-Tab para o Excel
		Sleep 500               ; Espera carregar
		Send "0"                ; Digita o valor padrão
		
		Sleep 10000             ; Esperar 10seg devido a lerdeza descomunal da Sonner
		
		Send "{DOWN}"           ; Desce para a próxima célula
		Send "!{TAB}"           ; Volta para o navegador
		Sleep 800               ; Espera carregar
	}
}

F10::ExitApp	                ; Encerra o script.