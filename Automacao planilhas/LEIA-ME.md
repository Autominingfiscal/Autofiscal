# Automação MVV em Python

Mesma cadeia de sempre, agora fora do Excel:

```
EXPEDIÇÃO  →  REMESSA  →  FATURAMENTO
              (+ PDFs da rede → chave de acesso)
```

A diferença prática é que não há mais macro dentro da planilha. As planilhas
voltam a ser `.xlsx` comuns e quem mexe nelas são os scripts desta pasta.
Cada rotina continua gravando uma cópia em `Backup` antes de tocar em
qualquer coisa.

| Atalho | Script | O que faz |
|---|---|---|
| `1 - atualizar remessas.bat` | `atualizar_remessas.py` | Lança as notas da expedição na aba Remessas Porto |
| `2 - preencher chaves.bat` | `preencher_chaves.py` | Lê os PDFs da rede e preenche a coluna CHAVE ACESSO |
| `3 - atualizar faturamento.bat` | `atualizar_faturamento.py` | Lança as remessas na aba do mês do faturamento |
| `0 - rodar tudo.bat` | `rodar_tudo.py` | Roda as três na ordem |

---

## Instalação (uma vez só)

**1. Conferir o Python**

Abra o Prompt de Comando e digite:

```
python --version
```

Precisa responder 3.8 ou mais novo. Se disser que não reconheceu o comando,
o Python não está no PATH — nesse caso use `py --version`, e troque `python`
por `py` dentro dos arquivos `.bat`.

**2. Instalar a única biblioteca necessária**

```
pip install openpyxl
```

Só isso. A leitura dos PDFs é feita com o `zlib`, que já vem junto com o
Python — não há mais nada para instalar. Se a empresa bloquear o download,
peça ao TI o pacote `openpyxl`; se faltar, os scripts avisam com essa mesma
instrução em vez de dar erro feio.

**3. Apontar os caminhos**

Rode qualquer um dos atalhos uma vez. Na primeira vez ele cria o arquivo
`config.ini` nesta pasta e para, pedindo para você preenchê-lo. Abra no
Bloco de Notas e ajuste:

```ini
[arquivos]
expedicao = \\servidor\pasta\EXPED_CONCENTRADO - 7 embarque.xlsx
remessa   = \\servidor\pasta\REMESSA FORM. LOTE MVV EMBARQUE 007 2026.xlsx
faturamento = \\servidor\pasta\FATURAMENTO 2026.xlsx
pdfs = \\servidor\pasta\PDFs
incluir_subpastas = SIM
```

Use sempre o caminho de rede (`\\servidor\pasta\...`), nunca a letra de
unidade mapeada — letra mapeada muda de máquina para máquina.

Mais abaixo no mesmo arquivo ficam o ano do embarque, o valor unitário padrão, o
CFOP, a operação e a transportadora padrão. **Mudou o preço do frete ou
virou o ano? Ajuste aí, não no código.**

---

## Uso no dia a dia

Duplo clique em `0 - rodar tudo.bat`. Uma janela preta abre, mostra o que
está fazendo e fecha quando você aperta Enter. Se preferir rodar uma etapa
de cada vez, use os atalhos numerados.

**Feche as planilhas no Excel antes de rodar.** Arquivo aberto não pode ser
gravado, e o script avisa isso em português em vez de travar.

---

## Parte 1 — Remessas Porto

Lê a planilha de expedição e reescreve a área de dados da aba Remessas
Porto, com os subtotais de cada dia e o rodapé fechando certo.

Pode rodar todo dia. Ela não "acrescenta" — reconcilia a aba inteira com a
expedição, então nota corrigida na expedição aparece corrigida aqui também.

**As chaves de acesso e os valores unitários já digitados são preservados**,
amarrados pelo número da nota. Rodar de novo não apaga o que você já preencheu.

**VL. UNITÁRIO varia por carregamento:** o `valor_unitario` do config.ini é só
o padrão que entra nas notas **novas**. Digite o valor certo na coluna G da
nota na própria planilha — o script mantém. Ao terminar, ele lista as notas
novas que ficaram com o valor padrão, para você conferir.

Os valores são gravados como número em Real (formato `R$ 1.234,56`, o mesmo
da planilha), então o Excel em português soma e calcula normalmente. Pode
digitar `15,20` ou até colar `R$ 15,20` como texto: o script converte para
número na próxima rodada. No config.ini também vale `valor_unitario = 14,80`.

Se você mudar o VL. UNITÁRIO de uma nota que **já estava no faturamento**, a
etapa do faturamento atualiza o VALOR dessa nota (só a coluna VALOR) e mostra
quantas mudaram.

Se a área de dados ficar pequena para o número de notas, ela cria as linhas
que faltam dentro da área, e as fórmulas do rodapé acompanham.

**Detalhe do porte:** a coluna PESO_LIQUIDO da expedição é uma fórmula, e o
arquivo não guarda o resultado dela. Fora do Excel não há como ler esse
valor pronto, então o script calcula: `PESO_SAÍDA − PESO_ENTRADA`, que é
exatamente o que a fórmula faz. Linha sem os dois pesos não é lançada e
aparece no relatório final.

---

## Parte 2 — Chaves de acesso a partir dos PDFs

Percorre as linhas da Remessas Porto que estão **sem chave**, procura na
pasta o PDF daquela nota, lê a chave de dentro do arquivo e grava no formato
de sempre (11 grupos de 4 dígitos). Linha que já tem chave não é tocada.

**Como ela acha o PDF de cada nota:** pelo primeiro número de 4 dígitos ou
mais no nome do arquivo. Funciona com `NF 19007 R. PORTO 19612.pdf` e também
com `19007.pdf`. Arquivo sem número no nome é ignorado, e só `.pdf` conta.

Desce nas subpastas em qualquer profundidade. Para varrer só a pasta
indicada, coloque `incluir_subpastas = NAO` no `config.ini`.

### As duas conferências

Nenhuma chave é gravada sem passar nas duas:

1. **Dígito verificador** da própria chave (módulo 11).
2. **O número da nota vem embutido na chave**, nas posições 26 a 34, e tem
   que bater com o número da nota daquela linha.

Ou seja: se o PDF errado estiver na pasta com o nome trocado, o script
percebe e **não grava** — lista a pendência no relatório para você conferir
na mão. Se o mesmo número aparecer em mais de uma subpasta, ele testa os
arquivos um por um e fica com o primeiro que passar nas duas conferências.

### Testar um PDF isolado

No Prompt de Comando, dentro desta pasta:

```
python preencher_chaves.py "\\servidor\pasta\NF 19007.pdf" 19007
```

Mostra a chave lida e a nota que está dentro dela. Útil quando uma nota
específica não preenche.

### Se uma nota não preencher

O relatório diz o motivo de cada uma:

- **PDF não encontrado** — o arquivo não está na árvore varrida, ou o nome
  não tem o número da nota.
- **não achei texto no PDF** — o PDF é digitalizado (imagem), não tem texto
  para extrair. Nesse caso precisaria de OCR; nem a macro nem o script
  resolvem isso.
- **a nota dentro da chave é X, esperava Y** — o PDF da pasta não é o
  daquela nota.

---

## Parte 3 — Faturamento

Lê a aba Remessas Porto e lança cada nota na aba do mês correspondente
(`09-26`, `10-26`, ...), determinada pela data da nota. Preenche 8 colunas:
DATA, Nº NF, VALOR, UF, CFOP, OPERAÇÃO, OUTRAS RECEITAS/SAÍDAS e KG/DIA.

Para cada nota, uma de três coisas acontece:

- **já lançada e completa** → não mexe;
- **a linha existe com só o Nº NF digitado** → completa a linha ali mesmo;
- **não existe** → insere na posição certa pela data.

Rodar de novo não duplica nada.

### O cuidado com o bloco de resumo

As fórmulas do resumo somam uma faixa fixa de linhas. Dentro do Excel, quem
esticava essas faixas ao inserir linhas era o próprio Excel. **Fora dele
isso não acontece sozinho** — o openpyxl empurra as células e deixa o texto
das fórmulas como estava, o que daria totais errados em silêncio.

Por isso este pacote reescreve as fórmulas ele mesmo, seguindo a regra do
Excel: inserindo N linhas na posição P, toda referência de linha ≥ P anda N
para baixo, nas duas pontas de cada faixa. Isso vale para as fórmulas de
todas as abas, para os intervalos mesclados, para a formatação condicional e
para os nomes definidos.

No fim, a rotina ainda confere se todas as faixas do resumo realmente cobrem
os dados e **avisa no relatório** se alguma ficou para trás, em vez de
deixar o total errado calado.

### Uma limitação que deixou de existir

A macro inseria as notas novas sempre no fim do mês, então nota retroativa
(um dia esquecido, lançado depois) deixava a coluna de datas fora de ordem.
Aqui a nota entra na posição certa pela data.

### Se um mês não tiver aba

Nota cuja data caia num mês sem aba correspondente não é lançada, e aparece
no relatório. O script não cria abas novas.

---

## O que muda em relação à macro

**A favor:** não precisa mais salvar as planilhas como `.xlsm` nem importar
módulo nenhum; a configuração é um arquivo de texto em vez de uma aba
escondida; a ordem de datas no faturamento fica certa mesmo com nota
retroativa; e dá para agendar no Agendador de Tarefas do Windows, coisa que
macro dentro de planilha não faz sozinha.

**Contra:** o Excel precisa estar fechado na hora de rodar. E o
`FATURAMENTO 2026.xlsx`, ao ser regravado fora do Excel, perde as
**configurações de impressão** de cada aba (margens, área de impressão,
escala). Nada de conteúdo, fórmula, formatação, comentário ou aba se perde —
isso foi conferido arquivo a arquivo. Mas se alguém imprime essas abas com
um ajuste específico, vale reconfigurar uma vez depois da primeira rodada.
A planilha de remessa não tem esse problema.

---

## Se der problema

**"Não consegui salvar"** — a planilha está aberta no Excel. Feche e rode
de novo.

**"Não achei a planilha de..."** — confira o caminho no `config.ini`. Se o
arquivo mudou de nome ou de pasta, o caminho antigo não vale mais.

**"Não encontrei a linha de rodapé"** — a rotina procura o texto
`VOLUME EMBARCADO` na coluna D para saber onde a área de dados termina. Se
essa linha foi renomeada ou apagada, ela para sem mexer em nada.

**"Não achei o bloco de resumo"** — a rotina procura `SUB-TOTAL` na coluna B
da aba do mês. Sem isso ela não lança nada naquela aba e avisa.

**Deu erro no meio** — a cópia da pasta `Backup` é a rede de proteção:
é só copiar de volta por cima.

**Faltou o openpyxl** — `pip install openpyxl` no Prompt de Comando.

---

## Como isto foi conferido

**Parte 1.** Rodada sobre a expedição real do 7º embarque contra a remessa
que a equipe já tinha preenchido à mão: as mesmas 327 notas, nos mesmos 12
blocos diários, nas mesmas linhas. **5.600 campos comparados, nenhuma
divergência.** As 253 chaves de acesso que já estavam lá foram devolvidas
para as notas certas.

**Parte 2.** O dígito verificador foi testado contra as **253 chaves reais**
da planilha: aprovou as 253, e o número da nota embutido bateu com a coluna
NF nas 253. Chave adulterada em um único dígito foi reprovada nas 43
posições testadas.

A extração recuperou a chave correta do PDF de teste e gravou no formato de
11 grupos. A varredura foi testada numa árvore montada de propósito, com
pasta vazia, arquivo que não é PDF, arquivo sem número no nome, a mesma nota
repetida em duas subpastas, um PDF com o nome de outra nota e um PDF
corrompido: achou os arquivos na árvore toda, recusou o de nome trocado com
a mensagem certa, ignorou o que tinha de ignorar e não travou no corrompido.

**Parte 3.** A premissa central — que fora do Excel as faixas do resumo
precisam ser reescritas à mão — foi testada nos dois casos: com folga dentro
da faixa e sem folga nenhuma.

Rodada sobre o faturamento real: completou as 67 linhas que tinham só o
número, e o subtotal de R.PORTO de 09-26 fechou em **R$ 161.730.256,00 e
10.927.720 kg** — exatamente o total da planilha de remessa, conferindo os
dois arquivos ponta a ponta. As outras 12 abas não mudaram nenhuma célula.

Depois, num cenário montado com 40 lançamentos removidos e a faixa do resumo
encolhida junto (sem folga nenhuma), a rotina inseriu as notas de volta, o
resumo esticou sozinho até cobrir a última linha de dados, o subtotal fechou
no mesmo valor e **a coluna de datas ficou sem nenhuma inversão** nas 338
linhas com data. Rodando uma segunda vez, nada foi duplicado: 327 notas, 327
já lançadas.

Também foram testados o mês com aba vazia (as notas entram a partir da
primeira linha, com data formatada como data) e a nota retroativa num mês
cheio de 453 lançamentos, que entrou na posição certa pela data.
