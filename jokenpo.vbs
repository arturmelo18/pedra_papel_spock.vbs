dim computador, jogador
call jokenpo
sub jokenpo()
    randomize(second(time))
    computador=int(rnd * 3) + 1
    jogador=cint(InputBox("[1]Pedra" + vbnewline &_
                          "[2]Papel" + vbnewline &_
                          "[3]Tesoura", "JOKENPO"))
    Select case jogador
        case 1:
            if computador = jogador then
                msgbox("Empate!"), vbinformation + vbokonly, "ATENCAO"
            elseif computador = 2 then
                msgbox("Derrota!"), vbinformation + vbokonly, "ATENCAO"
            else
                msgbox("Vitória!"), vbinformation + vbokonly, "ATENCAO"
            end if
        case 2:
            if computador = jogador then
                msgbox("Empate!"), vbinformation + vbokonly, "ATENCAO"
            elseif computador = 3 then
                msgbox("Derrota!"), vbinformation + vbokonly, "ATENCAO"
            else
                msgbox("Vitoria!"), vbinformation + vbokonly, "ATENCAO"
            end if
        case 3:
            if computador = jogador then
                msgbox("Empate!"), vbinformation + vbokonly, "ATENCAO"
            elseif computador = 1 then
                msgbox("Derrota!"), vbinformation + vbokonly, "ATENCAO"
            else
                msgbox("Vitoria!"), vbinformation + vbokonly, "ATENCAO"
            end if
        case else
            msgbox("Opcao invalida!"), vbExclamation + vbokonly, "ATENCAO"
            call jokenpo
    end Select
    jogador=msgbox("Voce quer jogar de novo?", vbQuestion + vbyesno, "FIM DO JOGO")
    if jogador=vbyes then
        call jokenpo
    else
        Wscript.quit
    end if
end sub