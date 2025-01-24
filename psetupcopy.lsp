(defun c:SETPSETUP ()
  (vl-load-com)

  (setq layouts (layoutlist))
  (setq acadObj (vlax-get-acad-object))
  (setq doc (vla-get-activedocument acadObj))

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
  (princ)
)

(princ "\nDigite SETPSETUP para executar o comando em todos os layouts.")
(princ)