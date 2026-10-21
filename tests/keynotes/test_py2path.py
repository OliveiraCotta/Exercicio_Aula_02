"""Simula o IronPython (texto UTF-16 por dentro, PY2=True): emoji = 2 unidades."""
import ast, sys
import os
HERE = os.path.dirname(os.path.abspath(__file__))
S = (sys.argv[1] if len(sys.argv) > 1 else "") or os.path.join(HERE, "..", "..", "OCA01.extension", "OCA.tab", "Keynote.panel", u"02_Atualização Descrição por Keynotes.pushbutton", "script.py"); tree = ast.parse(open(S, encoding="utf8").read())
W = {"TXT_BOM", "_txt_char_ok", "txt_encode", "_utf16le_units", "txt_decode", "validate_txt_bytes",
     "detect_txt_encoding", "TXT_ENCODING_LABEL"}
nodes = [n for n in tree.body if ({n.name} if isinstance(n, ast.FunctionDef) else
         {t.id for t in getattr(n, "targets", []) if isinstance(t, ast.Name)}) & W]
g = {"unichr": chr, "PY2": True}; exec(compile(ast.Module(body=nodes, type_ignores=[]), S, "exec"), g)
narrow = "A\tCaixa " + chr(0xD83D) + chr(0xDE00) + " ç\t\r\n"
b, bad = g["txt_encode"](narrow)
ok1 = not bad and bytes(b) == b"\xff\xfe" + "A\tCaixa \U0001F600 ç\t\r\n".encode("utf-16-le")
ok2 = g["txt_decode"](b) == narrow and g["validate_txt_bytes"](b, 1, narrow) == []
lone = "A\t" + chr(0xD83D) + "x\t\r\n"
ok3 = len(g["txt_encode"](lone)[1]) == 1
for name, ok in (("par substituto do IronPython vira os bytes certos", ok1),
                 ("leitura de volta igual + validação ok (caminho PY2)", ok2),
                 ("metade solta bloqueada (caminho PY2)", ok3)):
    print(("PASS " if ok else "FAIL ") + name)
