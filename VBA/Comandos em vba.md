
'definir um valor para a celula em [[vba/VBA]]

ActiveSheet.Range("a1").Value = "Luciano"

'selecionar uma celula especifica'

ActiveSheet.Range("a1").Select
ou

```Dim Plan1 as Worksheet
	set plan1 = sheets(Planilha1)
	Plan1.Cells(1,1).value = "Olá"
```


'seleciona a planilha 1'
sheets(1).select

'quando a planilha esta setadata e conte as colunas e recupere os valores usados
ultimaLinha = plan1.usedRange.rows.count

'copiar da plan 2 para a plan 1
plan2.range("a1:d" & UltimaLinha).copy
plan1.paste

'transfêrencia por igualdade'
plan21.range("a1:d" & UltimaLinha).value = plan?.range"a1:d" & UltimaLinha).value
'vai na ultima linha da planilha conta as colunas (Rows.count, 1 ) da coluna 1 e sobre a o começo e pula para a linha abaixo "+1"
ultimaLinha = Cells(Rows.Count, 1).End(xlUp).Row + 1

Mod comando para saber se um numero é PAR
If Target.Value Mod 2 = 0 Then







