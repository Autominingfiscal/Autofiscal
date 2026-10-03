# Novidades do Autofiscal

Cada versão nova entra no topo. A versão do programa fica em `painel/__init__.py`
(`VERSAO`) e aparece no rodapé do painel, no nome do instalador e em
*Propriedades > Detalhes* do `Autofiscal.exe`.

Como numerar (`MAIOR.MENOR.CORREÇÃO`):

- **CORREÇÃO** (1.1.**1**): conserto de erro, sem mudar como se usa.
- **MENOR** (1.**2**.0): ferramenta nova ou função nova.
- **MAIOR** (**2**.0.0): mudança que obriga a refazer configurações ou muda o jeito de usar.

---

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
