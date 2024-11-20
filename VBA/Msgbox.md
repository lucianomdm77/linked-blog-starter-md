# Quebra de Linha em {{{}}}

em [[vba/VBA]]
## vbCrLf ou Chr(13) ou Chr(10)

- Msgbox "comando de quebra de linha" & vbCrLf & "continua na segunda linha" &  vbCrLf & "!"
- Msgbox "comando de quebra de linha" & Chr(13) & "continua na segunda linha"  Chr(13) &  Chr(13) "!"
## vbCritical

- MsgBox "atenção! Erro ao executar", vbCritical, "Alerta"