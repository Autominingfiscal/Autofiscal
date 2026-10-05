# Novidades do Autofiscal

Cada versão nova entra no topo. A versão do programa fica em `painel/__init__.py`
(`VERSAO`) e aparece no rodapé do painel, no nome do instalador e em
*Propriedades > Detalhes* do `Autofiscal.exe`.

Como numerar (`MAIOR.MENOR.CORREÇÃO`):

- **CORREÇÃO** (1.1.**1**): conserto de erro, sem mudar como se usa.
- **MENOR** (1.**2**.0): ferramenta nova ou função nova.
- **MAIOR** (**2**.0.0): mudança que obriga a refazer configurações ou muda o jeito de usar.

---

## 1.5.0 — 05/10/2026

### Novo
- **Lançador da expedição → Lacres do turno** (vespertino ou matutino): com a
  EXPED_CONCENTRADO aberta, pede o par de lacres de cada motorista da tabela do
  turno que ainda não tem, confere (dois números, nenhum já usado em outra
  linha ou outro dia) e grava "9999 - 1265" na coluna CÓDIGO LACRE. Daí o
  Lançador leva o lacre para a expedição e para as OBSERVAÇÕES ADICIONAIS.
- Gera as **etiquetas** dos lacres, 15 por folha A4 com linha de corte (nome,
  cavalo, reboque, modelo, transportadora, lacres, turno e data), e abre no
  navegador para imprimir. Rodar de novo com todos preenchidos só reimprime.
- Motorista que já passou pela balança ganha o lacre também na linha da expedição.

## 1.4.1 — 05/10/2026

### Mudou
- **E-mails do faturamento** passa a funcionar com o **novo Outlook**, que não
  aceita automação. Botão *Preparar e-mails*: para cada e-mail, abre um e-mail
  novo já com destinatários, assunto e texto, e uma pasta só com os anexos
  dele. É só arrastar os arquivos (Ctrl+A), enviar e responder no painel para
  preparar o próximo. Os anexos não vão sozinhos: nenhum programa de e-mail
  aceita anexo por link.
- Saíram do painel os botões do Outlook clássico, que no PC da empresa podiam
  abrir o Outlook 2016 (que não funciona lá). Continuam na linha de comando
  (`--modo abrir` e `--modo enviar`).

## 1.4.0 — 05/10/2026

### Novo
- Ferramenta **E-mails do faturamento**: monta no Outlook os dois e-mails do
  dia, "FATURAMENTO / RELATÓRIO DE PESAGENS - <data>". O primeiro leva todos os
  XML da pasta do dia; o segundo, as NF, os tickets de pesagem e o relatório.
  Lê também as subpastas que o Arquivador cria e deixa de fora o `Backup`, os
  tickets assinados e os pré-cálculos. Mantém a assinatura do Outlook.
- Três botões: *Conferir anexos* (só lista), *Abrir e-mails para conferir*
  (você revisa e clica em Enviar) e *Enviar direto* (pede confirmação e não
  envia e-mail sem destinatário, sem anexo ou grande demais).
- Destinatários de cada e-mail em Configurações (`email_config.ini`).
- Precisa do Outlook clássico: o "novo Outlook" não aceita automação.

## 1.3.0 — 05/10/2026

### Novo
- **Lançador da expedição** preenche a coluna **OBSERVAÇÕES ADICIONAIS** (X)
  de cada caminhão: `TICKET N°: … / PLACA CAVALO: … / PLACA REBOQUE: … / LACRES N°: …`,
  com os dados da própria linha. Se o lacre ou uma placa for digitado ou
  corrigido depois, o texto se atualiza sozinho no ciclo seguinte (abas de hoje
  e de ontem). Texto escrito à mão que não comece com "TICKET N°:" não é trocado.

## 1.2.0 — 04/10/2026

### Novo
- Ferramenta **Inspecionar janela**: mostra o que o Windows enxerga dentro da
  janela de outro programa (campos, botões, rótulos), para saber se ele pode ser
  automatizado pelos campos da tela. É o primeiro passo para automatizar a
  emissão de notas no Datasul. Só lê: não clica nem digita nada. O arquivo
  gerado não leva o conteúdo dos campos de digitação, só o tamanho.
- *Listar janelas abertas* indica se uma janela vem de outro computador
  (Citrix, Área de Trabalho Remota), caso em que os campos não são visíveis.
- Biblioteca nova no instalador: pywinauto.

## 1.1.0 — 03/10/2026

### Novo
- **Instalador para Windows** (`Autofiscal-1.1.0-instalador.exe`): instala o
  Autofiscal em computadores **sem Python**. As bibliotecas (openpyxl, pypdf,
  pillow, pywin32) já vêm junto. Não pede administrador: instala na pasta do
  usuário e cria atalho no Menu Iniciar e, se quiser, na Área de Trabalho.
- Atualizar para uma versão nova **mantém os `.ini`** que o usuário já ajustou.
  Desinstalar também não apaga essas configurações.
- Para gerar o instalador, rode `distribuicao\Gerar instalador.bat`.
- Este arquivo de novidades e a numeração de versão.

### Mudou
- Instalado, o painel roda as ferramentas pelo `Autofiscal Executor.exe` em vez
  do `python.exe`. Rodando pelo código-fonte, continua tudo como antes.
- Instalado, o botão *Instalar* (pip) não aparece. As bibliotecas vêm no instalador.

## 1.0.0 — 02/10/2026

Primeira versão organizada do Autofiscal.

- **Painel** com todas as ferramentas numa janela só: executar, responder
  perguntas do script, editar as configurações `.ini` com *Procurar…*.
- Ferramentas: Remessas e Faturamento, Lançador da expedição, Painel da
  expedição, Conferência ISSQN, Conferência de retenções, Arquivar documentos
  do dia e Manutenção.
- **Pasta `comum`**: o código usado por mais de uma ferramenta fica num lugar só
  (chave da NF-e, números no jeito brasileiro, peso alvo e faixa de aceite, PDF).
- **Manutenção**: criar ferramenta nova a partir do modelo, exportar uma
  ferramenta para outro computador e rodar os testes.
- **Testes automáticos** de todas as ferramentas (`Rodar testes.bat`).
- Backups não se sobrescrevem no mesmo segundo e só os 30 mais recentes ficam.
- Correções: Arquivador sem sobrescrita silenciosa e com desfazer seguro,
  Lançador resistente a PDF ruim, leitura de PDF sem compressão, `.ini` não fica
  mais aberto depois de lido.
