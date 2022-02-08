INPUT:
1. AMOSTRAGENS.xlsx  [preenchido a partir do FT-03]
2. CLIENTE_e_PROJETO.txt [vem do comercial]
FT-14.xls

OUTPUT:
>> Aqui sairão os relatórios preenchidos no formato do TEMPLATE

PADRAO:
FT-23_rev 10 TEMPLATE.docx




1. Reservar espaço na FT-25 (Preencher as linhas para as amostragens a serem registradas)

2. Digitar as informações das FT-03 num excel, e colocar a FT-14 na mesma pasta.

3. Dar dois cliques em um arquivo .bat

----------------------------------------------------------------------------------------------------------------------

Ajustes:
Na hora de emitir a discrepância, ir até 
"INSTRUÇÕES DE PREENCHIMENTO DA CADEIA DE CUSTÓDIA" porque o lab pode ser outro "GÖRTLER" 

Na hora de emitir a discrepância, dependendo do Lab, pode haver outras combinação de parâmetros:

SVOC/TPH/PCB - 3*500ml

Na hora de emitir a discrepância, considerar os pepinos, pensar em algo robusto porque existem revisões e revisões

Comparar se os resultados dos parâmetros do FT-03 estão na "LQ e faixa de trabalho" certificados. Adcionar (há como?) aba no excel "1. AMOSTRAGENS.xlsx" com os "LQ e faixa de trabalho". Se estiver fora do LQ, emitir como > ou <. Botar formatação condicional avisando que isso aconteceu.

talvez remover lista do template, mas manter no excel

Se método de amostragem for "Baixa vazão - ABNT NBR 15.847:2010", "Vazão na coleta" nao pode ser maior que a "Vazão de purga"

No "2. CLIENTE_e_PROJETO.txt" (Pode virar uma aba de "1. AMOSTRAGENS"?) adicionar campo de escolha de assinaturas. A partir desse campo, substituir o .jpg em /media (com mesma razão de pixeis)

Criar e deletar o tmp em algum local não visualizável, talvez no "OUTPUT"

Correção automática se preenchido o nível do óleo (Falar Judite).
Se baixa vazão, "Nível inicial da amostragem" = "Nível final da purga"
Se purga mínima, "Nível inicial da amostragem" >= "Nível final da purga"

