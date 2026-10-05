# Autofiscal

Aplicação desktop para Windows que reúne, numa janela só, as ferramentas de
automação do setor fiscal: lançamento da expedição, remessas e faturamento,
conferências de impostos e arquivamento de documentos.

Cada ferramenta é um script Python. O painel mostra os botões de cada uma,
o que o script está fazendo, as perguntas dele e as configurações, sem precisar
abrir o Prompt de Comando.

![Versão](https://img.shields.io/badge/vers%C3%A3o-1.5.0-2563eb)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078d4)
![Python](https://img.shields.io/badge/Python-3.14-3776ab)

---

## Instalar (para quem só vai usar)

**Não precisa de Python.**

1. Baixe o `Autofiscal-<versão>-instalador.exe` em **[Releases](../../releases)**.
2. Dê dois cliques nele. Não pede administrador.
3. Abra o **Autofiscal** pelo Menu Iniciar ou pelo atalho na Área de Trabalho.

O programa é instalado na pasta do usuário (`%LOCALAPPDATA%\Programs\Autofiscal`).
Ao atualizar para uma versão nova, as configurações (`.ini`) que você já
ajustou são mantidas.

**Requisitos:** Windows 10 ou 11, 64 bits. O *Lançador da expedição* precisa do
Excel instalado. As outras ferramentas leem e gravam as planilhas sem ele.

---

## Ferramentas

| Ferramenta | O que faz |
|---|---|
| **Remessas e Faturamento** | Lança as notas da expedição na aba Remessas Porto, preenche as chaves de acesso a partir dos PDFs e alimenta a planilha de FATURAMENTO. |
| **Lançador da expedição** | Acompanha as pastas da balança e lança na planilha aberta no Excel a tara, o peso de saída (com a cor da faixa de aceite) e o número e a hora da nota. Monta a coluna OBSERVAÇÕES ADICIONAIS com ticket, placas e lacres de cada caminhão. Registra o par de lacres de cada motorista do turno e imprime as etiquetas (15 por folha). |
| **Painel da expedição** | Mostra o andamento da expedição no navegador: toneladas, faixa de aceite e caminhões na mina. Não usa a rede. |
| **Conferência ISSQN** | Confere a planilha de retenções de ISSQN com a conferência de retenções e com a pasta das notas. |
| **Conferência de retenções** | Compara, nota a nota, os impostos retidos do planilhão com a conferência de retenções e gera um relatório em Excel. |
| **Arquivar documentos do dia** | Distribui os documentos soltos da pasta do dia nas pastas das notas fiscais (NF, ticket, pré-cálculo…). Tem simulação e desfazer. |
| **E-mails do faturamento** | Monta no Outlook os dois e-mails do dia: um com os XML do faturamento e outro com as NF, os tickets e o relatório de pesagem. Funciona com o novo Outlook: abre cada e-mail já preenchido e uma pasta só com os anexos dele. |
| **Inspecionar janela** | Mostra o que o Windows enxerga dentro da janela de outro programa (ex.: Datasul), para saber se dá para automatizá-lo pelos campos da tela. Só lê, e não grava o conteúdo dos campos. |
| **Manutenção** | Cria uma ferramenta nova a partir do modelo, exporta uma ferramenta para outro computador e roda os testes. |

As ferramentas que mexem em planilhas guardam uma cópia em `Backup` antes de
gravar. Ficam as 30 cópias mais recentes de cada planilha.

---

## Desenvolvimento

### Rodar pelo código-fonte

Precisa de Python 3 e das bibliotecas:

```
pip install openpyxl pypdf pillow pywin32 pywinauto
```

Depois, dois cliques em **`Abrir Autofiscal.bat`** (ou em `Autofiscal.pyw`).
Cada ferramenta também roda sozinha, pelos `.bat` da pasta dela.

### Testes

```
python -m unittest discover -s tests -t tests
```

Ou dois cliques em **`Rodar testes.bat`**. Rode os testes depois de qualquer
mudança, principalmente na pasta `comum`, que todas as ferramentas usam.

### Estrutura

```
Autofiscal.pyw              abre o painel
painel/                     a janela (tkinter) e o executor das ferramentas
comum/                      código usado por mais de uma ferramenta
<Ferramenta>/               uma pasta por ferramenta, com ferramenta.json
_modelo_ferramenta/         modelo usado por Manutenção > Nova ferramenta
tests/                      testes automáticos (unittest)
distribuicao/               geração do .exe e do instalador
CHANGELOG.md                o que mudou em cada versão
```

Toda subpasta com um arquivo `ferramenta.json` aparece no painel como uma
ferramenta. O `ferramenta.json` define o nome, os botões, o script de cada
botão e as perguntas (pasta, arquivo ou texto) feitas antes de rodar.

O passo a passo para criar uma ferramenta, o formato do `ferramenta.json` e o
que tem na pasta `comum` estão em
**[COMO ADICIONAR UMA FERRAMENTA.md](COMO%20ADICIONAR%20UMA%20FERRAMENTA.md)**.

### Gerar o instalador

Só o computador que **gera** o instalador precisa disto:

1. Python 3 e `pip install pyinstaller openpyxl pypdf pillow pywin32 pywinauto`
2. [Inno Setup 6](https://jrsoftware.org/isinfo.php)

Depois:

1. Suba a versão em `painel/__init__.py` (`VERSAO`) e registre as mudanças no
   [`CHANGELOG.md`](CHANGELOG.md).
2. Rode os testes.
3. Dois cliques em **`distribuicao\Gerar instalador.bat`**.

O instalador sai em `distribuicao\saida\Autofiscal-<versão>-instalador.exe`.
Ele não entra no git. Para distribuir, anexe-o a uma **Release** com a tag
`v<versão>` (por exemplo `v1.1.0`).

**Como o programa funciona sem Python:** o PyInstaller leva o interpretador e
as bibliotecas para a pasta `_internal`. O `Autofiscal.exe` abre a janela, e o
`Autofiscal Executor.exe` roda os scripts `.py` das ferramentas no lugar do
`python.exe` (veja `painel/executor.py`). As ferramentas continuam como
código-fonte ao lado do `.exe`.

### Versões

O número segue `MAIOR.MENOR.CORREÇÃO`:

- **CORREÇÃO** (1.1.**1**): conserto de erro, sem mudar como se usa.
- **MENOR** (1.**2**.0): ferramenta ou função nova.
- **MAIOR** (**2**.0.0): mudança que obriga a refazer configurações ou muda o jeito de usar.

O que mudou em cada versão está no [CHANGELOG.md](CHANGELOG.md).
