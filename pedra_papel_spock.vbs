Dim resp, comp

call jogo
sub jogo()
    randomize(second(time))
    comp=int(rnd*5)+1

    resp=int(inputbox("[1]Pedra"+vbnewline &_
                      "[2]Papel"+vbnewline &_
                      "[3]Tesosura"+vbnewline&_
                      "[4]Lagarto"+vbnewline&_
                      "[5]Spock"+vbnewline &_
                      "Escolha uma opção:", "JOKENPO"))
    Select case resp:
        case 1:
            if comp = resp then
                msgbox("Empate!"),vbquestion+vbokonly,"JOKENPO"
            elseif comp = 2 or comp = 5 then
                msgbox("Derrota!"),vbquestion+vbokonly,"JOKENPO"
            else
                msgbox("Vitória!"),vbquestion+vbokonly,"JOKENPO"
            end if
            call pergunta
        case 2:
            if comp = resp then
                msgbox("Empate!"),vbquestion+vbokonly,"JOKENPO"
            elseif comp = 3 or comp = 4 then
                msgbox("Derrota!"),vbquestion+vbokonly,"JOKENPO"
            else
                msgbox("Vitória!"),vbquestion+vbokonly,"JOKENPO"
            end if
            call pergunta
        case 3:
            if comp = resp then
                msgbox("Empate!"),vbquestion+vbokonly,"JOKENPO"
            elseif comp = 1 or comp = 5 then
                msgbox("Derrota!"),vbquestion+vbokonly,"JOKENPO"
            else
                msgbox("Vitória!"),vbquestion+vbokonly,"JOKENPO"
            end if
            call pergunta
        case 4:
            if comp = resp then
                msgbox("Empate!"),vbquestion+vbokonly,"JOKENPO"
            elseif comp = 1 or comp = 3 then
                msgbox("Derrota!"),vbquestion+vbokonly,"JOKENPO"
            else
                msgbox("Vitória!"),vbquestion+vbokonly,"JOKENPO"
            end if
            call pergunta
        case 5:
            if comp = resp then
                msgbox("Empate!"),vbquestion+vbokonly,"JOKENPO"
            elseif comp = 2 or comp = 4 then
                msgbox("Derrota!"),vbquestion+vbokonly,"JOKENPO"
            else
                msgbox("Vitória!"),vbquestion+vbokonly,"JOKENPO"
            end if
            call pergunta
        case else:
            msgbox("Opção inválida!"),vbexclamation + vbokonly,"ATENÇÃO"
            call jogo
    end Select
end sub

function pergunta()
    resp=msgbox("Você deseja jogar de novo?",vbquestion + vbyesno,"JOKENPO")
        if resp=vbYes then
            call jogo
        else
            Wscript.quit
        end if
end function