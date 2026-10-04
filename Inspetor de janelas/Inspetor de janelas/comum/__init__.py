"""Codigo compartilhado por todas as ferramentas do Autofiscal.

Tudo que mais de uma ferramenta precisa mora AQUI, uma vez so:

    caminhos   caminhos que funcionam para qualquer usuario do Windows
    numeros    numero escrito no jeito brasileiro ("R$ 1.234,56")
    texto      comparacao de textos sem acento/pontuacao (norm)
    arquivos   leitura de .ini e .txt em UTF-8 ou ANSI
    nfe        chave de acesso da NF-e (digito verificador, numero da nota)
    pesagem    peso bruto alvo por modelo e faixa de aceite do carregamento
    pdf        pecas de baixo nivel para ler texto de PDF sem biblioteca

Cada ferramenta acha esta pasta sozinha: na pasta Autofiscal (logo acima da
ferramenta) ou dentro da propria ferramenta, quando ela foi exportada para
rodar em outro computador (ferramenta Manutencao > Exportar).

Mudou algo aqui? Rode os testes ("Rodar testes.bat"): todas as ferramentas
usam este codigo. Nada aqui depende de biblioteca fora do Python.
"""

VERSAO = "1.0"
