# -*- coding: utf-8 -*-
"""Tokens do Design System OCA: o único lugar com cores, fontes e medidas.

Janelas WPF, relatórios HTML e planilhas Excel leem estes valores. Os comandos
nunca usam código hexadecimal: referenciam os pincéis "oca.*" (WPF), as
variáveis "--oca-*" (HTML) ou os estilos do módulo excel.

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3.
"""

# Marca: #2FBAA7 (cor do logo), igual nos dois temas: logo e botão principal.
# Accent: mesmo matiz com contraste >= 4,5:1 para texto, foco e seleção.
# FieldLine: borda de campo com contraste >= 3:1 sobre o fundo do campo.
LIGHT = {
    "Bg": "#F3F5F6",          # fundo da janela
    "Surface": "#FFFFFF",     # cabeçalho, rodapé, seções, cartões
    "Surface2": "#EEF1F2",    # cabeçalho de tabela, linha alternada
    "Field": "#FFFFFF",       # campos, listas, tabelas
    "Line": "#D5DBDF",        # separadores e contornos
    "FieldLine": "#808C94",   # borda de campo (3,4:1)
    "Ink": "#1D2428",         # texto
    "Ink2": "#5A666D",        # texto secundário
    "Accent": "#1D7266",      # foco, seleção, links
    "AccentInk": "#FFFFFF",   # texto/ícone sobre Accent
    "AccentSoft": "#E3F4F1",  # fundo de seleção e de mensagem informativa
    "Brand": "#2FBAA7",       # logo e botão principal
    "BrandInk": "#0E2226",    # texto do botão principal (6,8:1)
    "Ok": "#1D7438",
    "OkSoft": "#E5F3E9",
    "Warn": "#8A5100",
    "WarnSoft": "#FCF0DB",
    "Err": "#AE251D",
    "ErrSoft": "#FCE4E2",
    "Violet": "#6A55C2",      # identificador (troca de tipo)
}

# Tema escuro: opção A, grafite neutro (painel #33393F).
DARK = {
    "Bg": "#2A2F35",
    "Surface": "#33393F",
    "Surface2": "#2F343A",
    "Field": "#24292E",
    "Line": "#474F57",
    "FieldLine": "#7C8790",   # 4,0:1 sobre o campo, 3,2:1 sobre o painel
    "Ink": "#E7EBEE",
    "Ink2": "#A9B3BA",
    "Accent": "#3FC0AE",
    "AccentInk": "#0E2226",
    "AccentSoft": "#2D4346",  # texto Accent sobre ele: 4,7:1
    "Brand": "#2FBAA7",
    "BrandInk": "#0E2226",
    "Ok": "#73C98E",
    "OkSoft": "#344541",
    "Warn": "#EFAE55",
    "WarnSoft": "#464139",
    "Err": "#F28C84",
    "ErrSoft": "#423A3E",
    "Violet": "#B7A9F0",
}

# Raios (px): sem cantos vivos dentro das janelas.
RADIUS = {
    "control": 4,    # campos, listas, botões, etiquetas pequenas
    "group": 6,      # seções, mensagens, tabelas, cartões, dicas
    "panel": 8,      # painéis grandes e menus suspensos
}

# Larguras padrão das janelas (px): S formulários, M formulários longos,
# L grades (redimensionável).
SIZES = {"S": 460, "M": 580, "L": 1160}
LABEL_WIDTH = 120
FIELD_HEIGHT = 26
BUTTON_HEIGHT = 28

FONT = u"Segoe UI"
FONT_MONO = u"Consolas"

# Excel: sempre claro (planilhas são impressas e abertas fora do Revit).
# Valores ARGB, como o OOXML espera.
EXCEL = {
    "font": u"Segoe UI",
    "font_mono": u"Consolas",
    "size": 10,
    "ink": "FF1D2428",
    "ink2": "FF5A666D",
    "header_fill": "FF1D7266",   # branco sobre ele: 5,75:1
    "header_ink": "FFFFFFFF",
    "band": "FFF3F5F6",          # linha alternada
    "line": "FFD5DBDF",
    "brand": "FF2FBAA7",
    "warn_soft": "FFFCF0DB",
    "warn": "FF8A5100",
}
