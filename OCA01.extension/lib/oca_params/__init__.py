# -*- coding: utf-8 -*-
"""Parâmetros compartilhados: leitura do TXT, plano de vinculação e API do Revit.

Módulos:
    spfile     leitura do arquivo TXT de parâmetros compartilhados (somente leitura,
               Python puro: não importa a API do Revit)
    plan       plano de execução (o que será criado, ampliado ou bloqueado) a
               partir de dados simples (Python puro)
    revit_ops  tudo que usa a API do Revit: grupos de parâmetros por versão,
               categorias, vínculos existentes e a transação de vinculação

Usado pelo comando "Gerenciador de Parâmetros Compartilhados" (aba OCA).
Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3.
"""
