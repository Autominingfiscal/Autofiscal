# Autofiscal: como usar e como adicionar uma ferramenta

## Abrir

Dê dois cliques em **Abrir Autofiscal.bat**, ou direto em **Autofiscal.pyw**.
Para ter um atalho na Área de Trabalho, clique com o botão direito no `.bat`,
escolha *Enviar para > Área de trabalho (criar atalho)*.

Na janela:

- **Lista à esquerda:** as ferramentas. A bolinha verde indica que a ferramenta está rodando.
- **Aba Executar:** os botões de cada ação, um registro do que o programa está
  fazendo e um campo para responder quando ele pergunta alguma coisa.
- **Aba Configurações:** os campos do arquivo `.ini` da ferramenta, com o botão
  *Procurar…* para escolher pastas e arquivos. Ao lado de cada caminho aparece se
  ele foi encontrado neste computador. *Salvar alterações* grava só os campos
  alterados e mantém os comentários do arquivo.

Cada ferramenta continua funcionando sozinha, pelos `.bat` dela, como antes.

O que o painel lembra, como a última pasta escolhida, fica guardado por usuário
do Windows, em `%APPDATA%\Autofiscal`. Assim, um usuário não atrapalha o outro.

---

## Adicionar uma ferramenta nova

1. No painel, abra **Manutenção** e clique em **Nova ferramenta**. Digite o nome
   que vai aparecer no painel. A pasta é criada a partir de `_modelo_ferramenta`,
   já com um script de exemplo e o `ferramenta.json`.
2. Clique em **Recarregar ferramentas**. A ferramenta nova já aparece na lista.
3. Troque o `exemplo.py` pelo script da ferramenta e ajuste os botões no
   `ferramenta.json` (modelo abaixo). **Mantenha o bloco "comum" do começo do script.**
4. Em **Manutenção**, clique em **Rodar testes**.

Também dá para fazer à mão: crie uma subpasta, copie para ela o conteúdo de
`_modelo_ferramenta` e edite.

Pastas que começam com `_` são ignoradas. Por isso `_modelo_ferramenta` não aparece no painel.

---

## A pasta `comum`: código que todas as ferramentas usam

O que mais de uma ferramenta precisa fica **uma vez só** na pasta `comum`. Se
precisar disso numa ferramenta nova, use daqui. Não copie para dentro da ferramenta.

| Módulo | O que tem |
|---|---|
| `comum.caminhos` | `caminho_do_usuario` (`%OneDrive%`, caminho de outro usuário, relativo) e `existe` (aceita `*`) |
| `comum.numeros` | `numero_br`: `"R$ 1.234,56"`, `"14,80"`, `"(10,00)"` viram número; texto e vazio viram `None` |
| `comum.texto` | `norm` (compara cabeçalho, placa e modelo sem acento nem pontuação) e `sem_acento` |
| `comum.arquivos` | `ler_ini` e `ler_texto`: aceitam arquivo salvo em UTF-8 ou ANSI pelo Bloco de Notas |
| `comum.nfe` | chave de acesso: `dv_confere`, `nota_da_chave`, `modelo_da_chave`, `formatar` |
| `comum.pesagem` | `peso_alvo` por modelo e `faixa_de_aceite` (vermelho, amarelo, verde ou dentro) |
| `comum.pdf` | peças para ler texto de PDF sem biblioteca: `streams`, `descompactar`, `desescapar` |

Para usar, o script começa com o **bloco "comum"** (já vem no modelo). Ele
procura a pasta `comum` ao lado do script e, se não achar, na pasta Autofiscal,
logo acima. Depois é só importar:

```python
from comum.numeros import numero_br
from comum.caminhos import caminho_do_usuario
```

**Mudou algo na pasta `comum`?** Todas as ferramentas usam esse código. Rode os
testes antes de usar (**Manutenção > Rodar testes** ou `Rodar testes.bat`). Os
testes também acusam se alguma ferramenta voltar a ter uma cópia de uma função
que já está na `comum`.

---

## Usar uma ferramenta em outro computador

Algumas ferramentas rodam em outro PC, sem o resto do Autofiscal (por
exemplo, o Lançador no computador da balança).

1. Em **Manutenção**, clique em **Exportar ferramenta**.
2. Escolha a pasta da ferramenta e depois o destino (pendrive, rede ou a pasta do outro PC).
3. Ela é copiada **com a pasta `comum` dentro** e roda sozinha lá.

**Para atualizar** depois de uma mudança, exporte de novo para o mesmo destino:

- o código (`.py`, `.bat`, `.md`, `ferramenta.json`) é substituído pela versão nova;
- as configurações que já estiverem lá (`.ini` e outros `.json`) são **mantidas**;
- planilhas, PDFs, logs e a pasta `Backup` nunca são copiados.

### Modelo de `ferramenta.json`

```json
{
  "nome": "Nome que aparece no painel",
  "descricao": "Uma frase dizendo o que ela faz.",
  "ordem": 70,
  "config": "minha_config.ini",
  "ajuda": "LEIA-ME.md",
  "requer": [{"modulo": "openpyxl", "pacote": "openpyxl"}],
  "acoes": [
    {
      "rotulo": "Rodar",
      "script": "meu_script.py",
      "principal": true,
      "descricao": "O que este botão faz."
    },
    {
      "rotulo": "Rodar numa pasta",
      "script": "meu_script.py",
      "args": ["{pasta:Escolha a pasta}", "--opcao"],
      "descricao": "Pergunta a pasta antes de rodar."
    }
  ]
}
```

| Campo | Para que serve |
|---|---|
| `nome`, `descricao` | O que aparece na lista e no cartão da tela inicial. |
| `ordem` | Posição na lista (número menor fica em cima). |
| `config` | Arquivo `.ini` da ferramenta. Com ele aparece a aba Configurações. Opcional. |
| `ajuda` | Arquivo aberto pelo botão Ajuda. Opcional. |
| `requer` | Bibliotecas usadas. Se faltar alguma, o painel avisa e oferece o botão *Instalar*. |
| `campos` | Opcional. Força o tipo de um campo do `.ini`: `{"planilha": {"tipo": "texto", "ajuda": "..."}}`. Tipos: `pasta`, `arquivo`, `texto`, `simnao`. |
| `acoes` | Os botões da aba Executar. |

### Campos de cada ação

| Campo | Para que serve |
|---|---|
| `rotulo` | Texto do botão. |
| `script` | Arquivo `.py` dentro da pasta da ferramenta. |
| `args` | Argumentos passados ao script (opcional, veja abaixo). |
| `continuo` | `true` para programas que ficam rodando até alguém clicar em **Parar** (monitoramento, painel). |
| `confirmar` | Mensagem de "tem certeza?" antes de rodar. |
| `principal` | `true` deixa o botão azul. |
| `descricao` | Texto ao lado do botão. |

### Perguntas antes de rodar (dentro de `args`)

| Marcador | O painel abre |
|---|---|
| `{pasta:Título}` | janela para escolher uma pasta |
| `{arquivo:Título\|*.pdf *.xlsx}` | janela para escolher um arquivo (o filtro depois do `\|` é opcional) |
| `{texto:Pergunta}` | caixa para digitar um texto |

- **Lembrança:** o painel guarda a última resposta de cada marcador.
- **Resposta opcional:** com `?` depois do tipo, como `{texto?:...}`, a resposta pode ficar vazia.
- **Grupo de argumentos:** coloque um marcador opcional dentro de uma lista para que o grupo inteiro saia quando a resposta ficar vazia. Exemplo: `["--faixa", "{texto?:Faixa de notas}"]`.

### Regras para o script funcionar bem no painel

- **Saída na tela:** escreva com `print()`. Tudo aparece no registro. Linhas com
  `ERRO` ficam vermelhas e linhas com `ATENÇÃO` ficam amarelas.
- **Perguntas com `input()`:** funcionam. O usuário responde no campo de
  resposta, embaixo do registro.
- **Pausa no final:** não segure a janela com `input("Pressione Enter...")`. Use
  isto, que só pausa quando o script é aberto fora do painel:

  ```python
  if sys.stdin is not None and sys.stdin.isatty():
      input("Pressione Enter para fechar...")
  ```

- **Detectar o painel:** o painel define a variável de ambiente `AUTOFISCAL=1`, caso o script precise saber se está rodando dentro dele.
- **Caminhos:** use a função `caminho_do_usuario` (em `comum\caminhos.py`), para
  que os caminhos funcionem para qualquer usuário do Windows.

---

## Gerar o instalador (para computadores sem Python)

1. Neste computador, uma vez só: `pip install pyinstaller openpyxl pypdf pillow pywin32`
   e o [Inno Setup 6](https://jrsoftware.org/isinfo.php).
2. Suba a versão em `painel/__init__.py` (`VERSAO`) e conte o que mudou no
   `CHANGELOG.md`.
3. Rode os testes e depois dê dois cliques em `distribuicao\Gerar instalador.bat`.
4. O instalador sai em `distribuicao\saida\Autofiscal-<versão>-instalador.exe`.

Uma ferramenta nova entra no instalador sozinha. Se ela usar uma biblioteca nova,
basta essa biblioteca estar instalada no computador que gera o instalador: os
`import` dos scripts são encontrados automaticamente.

Os `.ini` que vão no instalador são os desta pasta. Na primeira instalação eles
viram as configurações iniciais. Nas atualizações, o instalador **não**
sobrescreve os `.ini` que já estão no outro computador.
