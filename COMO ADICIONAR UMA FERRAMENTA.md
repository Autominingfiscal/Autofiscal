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

1. Crie uma subpasta dentro de `Autofiscal`, por exemplo `Autofiscal\Nova ferramenta`.
2. Coloque nela o script `.py`.
3. Crie nela um arquivo **`ferramenta.json`** (modelo abaixo).
4. No painel, clique em **Recarregar ferramentas**.

Pastas que começam com `_` são ignoradas. Por isso `_modelo_ferramenta` não aparece no painel.

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
- **Caminhos:** use a função `caminho_do_usuario` (em `painel\caminhos.py`), para
  que os caminhos funcionem para qualquer usuário do Windows.
