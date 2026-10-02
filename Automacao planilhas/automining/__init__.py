"""Automacao MVV - rotinas da balanca em Python.

Porte dos modulos VBA (Mod_RemessasPorto, Mod_ChaveAcesso, Mod_Faturamento)
para Python + openpyxl, mantendo a mesma logica ja conferida contra os dados
reais.

A cadeia do dia a dia continua a mesma:

    EXPEDICAO  ->  REMESSA  ->  FATURAMENTO
                   (+ PDFs da rede -> chave de acesso)
"""

__version__ = "1.0"
