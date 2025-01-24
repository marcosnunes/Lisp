(defun c:SETPSETUP ()
  (vl-load-com)

  (setq layouts (layoutlist))
  (setq acadObj (vlax-get-acad-object))
  (setq doc (vla-get-activedocument acadObj))
  (setq setup_a3_found nil) ; Variável para verificar se o layout Setup_A3 foi encontrado

  ; Verifica se o layout Setup_A3 existe
  (foreach layout layouts
    (if (equal layout "Setup_A3")
      (setq setup_a3_found t)
    )
  )

  (if setup_a3_found ; Se o layout Setup_A3 foi encontrado, executa o restante do código
    (progn
      (foreach layout layouts
        (if (not (equal layout "Model"))
          (progn
            ; Obtém o objeto Layout para o layout atual
            (setq layoutObj (vla-item (vla-get-layouts doc) layout))

            ; Define a configuração de página para o layout atual usando o método CopyFrom
            (vl-catch-all-apply
              '(lambda ()
                (vla-copyfrom layoutObj (vla-item (vla-get-layouts doc) "Setup_A3"))
              )
            )
            ;Define o layout ativo
            (vla-put-activelayout doc layoutObj)

            (princ (strcat "\nConfiguração de página 'Setup_A3' aplicada no layout: " layout))
          )
        )
      )
      (princ "\nProcesso concluído em todos os layouts.")
    )
     (princ "\nLayout 'Setup_A3' não encontrado no desenho. Você precisa criar um layout e nomeá-lo como Setup_A3 e fazer as configurações de impressão nele para que sejam copiadas para os outros layouts. Operação cancelada") ;Mensagem se o layout não for encontrado
  )
  (princ)
)

(princ "\nDigite SETPSETUP para executar o comando em todos os layouts.")
(princ)
