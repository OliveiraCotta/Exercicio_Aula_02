"""Uso: python3 tests/keynotes/test_txt16.py [script.py] [TXT do escritório]
Testes offline do TXT de Keynote em UTF-16 LE com BOM e do backup em <nome>_obsoletos.
Extrai as funções do script por AST (não executa o comando). Referência dos bytes:
o codec 'utf-16-le' do CPython (independente do código testado)."""
import ast, os, sys, shutil, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fake_dotnet as shim
from System.IO import File
HERE = os.path.dirname(os.path.abspath(__file__))
S = (sys.argv[1] if len(sys.argv) > 1 else "") or os.path.join(
    HERE, "..", "..", "OCA01.extension", "OCA.tab", "Keynote.panel",
    u"02_Atualização Descrição por Keynotes.pushbutton", "script.py")
OFFICE = (sys.argv[2] if len(sys.argv) > 2 else "") or os.path.join(   # TXT real do escritório
    HERE, "fixtures", "OCA_XXX_Lista_de_Materiais_00_ANSI.txt")
tree = ast.parse(open(S, encoding="utf8").read())
WANT = {"TXT_BOM", "TXT_ENCODING_LABEL", "_TXT_HINT", "_txt_char_ok", "txt_encode",
        "_utf16le_units", "txt_decode", "detect_txt_encoding", "txt_char_report", "_bytes_of",
        "validate_txt_bytes", "OBSOLETE_SUFFIX", "obsolete_backup_path", "_io_message",
        "write_keynote_txt", "_no_tab", "to_unicode", "_txt_field"}
nodes = []
for n in tree.body:
    names = [n.name] if isinstance(n, (ast.FunctionDef, ast.ClassDef)) else \
        [t.id for t in getattr(n, "targets", []) if isinstance(t, ast.Name)]
    if set(names) & WANT:
        nodes.append(n)
class _Log:
    def __init__(self): self.lines = []
    def __getattr__(self, k): return lambda *a: self.lines.append((k,) + a)
from datetime import datetime
LOG = _Log()
g = {"os": os, "File": File, "datetime": datetime, "LOG": LOG, "unicode": str, "unichr": chr,
     "PY2": False}
exec(compile(ast.Module(body=nodes, type_ignores=[]), S, "exec"), g)
missing = WANT - set(g)
R = []
def check(name, cond, info=""):
    R.append((name, bool(cond)))
    print(("PASS " if cond else "FAIL ") + name + ("  " + str(info)[:160] if info else ""))
def ref(text):                       # bytes de referência: BOM FF FE + codec do CPython
    return list(b"\xff\xfe" + text.encode("utf-16-le"))
enc, dec, val, det = g["txt_encode"], g["txt_decode"], g["validate_txt_bytes"], g["detect_txt_encoding"]
LABEL = g["TXT_ENCODING_LABEL"]
check("rótulo da codificação = 'UTF-16 LE, com BOM'", LABEL == u"UTF-16 LE, com BOM")

print("\n# 1. Acentos e cedilha")
for word in (u"á", u"ã", u"é", u"í", u"ó", u"ú", u"ç", u"ÁÃÉÍÓÚÇ", u"Ação · Pé-direito · Coração · Ônibus · Ü"):
    b, bad = enc(word)
    check("%s -> bytes = referência, ida e volta" % word, b == ref(word) and not bad and dec(b) == word)
b, _ = enc(u"ç")
check("ç = FF FE E7 00 (UTF-16 LE com BOM, não BE)", b == [0xFF, 0xFE, 0xE7, 0x00], b)

print("\n# 2. Caracteres especiais e símbolos (todos representáveis em UTF-16)")
SPECIAL = u"€ “aspas” ‘x’ – — … • ™ ® © § ¶ ° ± × ÷ ² ³ µ ½ ‰ ≥ ≤ ≈ ≠ Ω → ∅ Ø ✓ 中文 😀  nbsp"
b, bad = enc(SPECIAL)
check("símbolos, ≥, Ω, →, CJK, emoji: bytes = referência", b == ref(SPECIAL) and not bad, bad)
check("símbolos: ida e volta idêntica", dec(b) == SPECIAL)
b, _ = enc(u"😀")
check("emoji vira par substituto D83D DE00", b == [0xFF, 0xFE, 0x3D, 0xD8, 0x00, 0xDE], b)

print("\n# 3. Não representáveis (listados, nada gravado)")
for ch, why in ((u"\x00", "NUL"), (u"\x07", "BEL"), (u"\x1f", "controle C0"), (u"\x7f", "DEL"),
                (u"\x85", "controle C1"), (u"﻿", "BOM no meio do texto"), (u"￾", "não-caractere"),
                (u"\ud800", "1ª metade solta"), (u"\udc00", "2ª metade solta")):
    b, bad = enc(u"A" + ch + u"B")
    check("bloqueia %s (U+%04X)" % (why, ord(ch)), len(bad) == 1 and bad[0][1] == ch)
rep = g["txt_char_report"]([(u"RE01", u"Texto\x07com controle", u"RE", u"aba MATERIAIS linha 7"),
                            (u"RE02", u"Texto ok ≥ 2", u"RE", u"aba MATERIAIS linha 8")])
check("relatório: só o controle, com código, keynote, origem e dica",
      len(rep) == 1 and rep[0]["code"] == u"U+0007" and rep[0]["key"] == u"RE01"
      and rep[0]["origin"] == u"aba MATERIAIS linha 7" and u"controle" in rep[0]["hint"], rep)

print("\n# 4. Quebras de linha e linhas vazias")
nt = g["_no_tab"]
check("quebra dentro da célula (CRLF, LF, CR) e TAB viram espaço",
      nt(u"a\r\nb\nc\rd\te") == u"a  b c d e")
good = u"A\tTexto ç\t\r\nA01\tSub ã\tA\r\n"
check("estrutura válida aceita (CRLF, 2 TAB por linha)", val(ref(good), 2, good) == [])
check("LF sem CR detectado", val(ref(u"A\tT\t\nB\tT\tA\n"), 2))
check("CR solto detectado", val(ref(u"A\tT\t\rB\tT\tA\r\n"), 2))
check("linha vazia no meio detectada", val(ref(u"A\tT\t\r\n\r\nB\tT\tA\r\n"), 2))
check("TAB a mais detectado", val(ref(u"A\tT\tx\t\r\n"), 1))
check("contagem de linhas divergente detectada", val(ref(good), 3))
check("texto lido diferente do gerado detectado", val(ref(good), 2, good.replace(u"ç", u"c")))

print("\n# 5. BOM e codificações (detecção e validação)")
t = u"PI\tPISO ç\t\r\n"
cases = [("UTF-16 LE, com BOM", list(b"\xff\xfe" + t.encode("utf-16-le"))),
         ("UTF-16 BE, com BOM", list(b"\xfe\xff" + t.encode("utf-16-be"))),
         ("UTF-16 LE, sem BOM", list(t.encode("utf-16-le"))),
         ("UTF-16 BE, sem BOM", list(t.encode("utf-16-be"))),
         ("UTF-8, com BOM", list(b"\xef\xbb\xbf" + t.encode("utf-8"))),
         ("UTF-8, sem BOM", list(t.encode("utf-8"))),
         ("ANSI (Windows-1252)", list(t.encode("cp1252"))),
         ("ASCII (compatível com ANSI)", list(b"PI\tPISO\t\r\n"))]
for name, data in cases:
    check("detecta %s" % name, det(data) == name, det(data))
    if name != "UTF-16 LE, com BOM":
        check("validação recusa %s" % name, val(data, 1) != [])
check("validação aceita só UTF-16 LE com BOM", val(cases[0][1], 1, t) == [])
check("tamanho ímpar (arquivo truncado) detectado", val(cases[0][1] + [0x41], 1))
check("NUL no arquivo detectado", val(list(b"\xff\xfe" + u"A\tB\x00\t\r\n".encode("utf-16-le")), 1))
check("metade de par solta no arquivo detectada",
      val([0xFF, 0xFE, 0x41, 0, 9, 0, 0x00, 0xD8, 9, 0, 13, 0, 10, 0], 1))

print("\n# 6. Gravação, backup e arquivos existentes em outras codificações")
import tempfile
d = os.path.join(tempfile.mkdtemp(), "ut16"); shutil.rmtree(d, ignore_errors=True); os.makedirs(d)
w = g["write_keynote_txt"]
p = os.path.join(d, u"OCA_123_Lista de Materiais_01.txt")
text1 = u"PI\tPISO\t\r\nPI01\tConcreto polido - ação “A” ≥ 7 cm\tPI\r\n"
bk, encn, prev = w(p, text1, 2)
raw = open(p, "rb").read()
check("1ª gravação: bytes no disco = FF FE + UTF-16 LE do texto", raw == b"\xff\xfe" + text1.encode("utf-16-le"))
check("1ª gravação: rótulo devolvido confere com os bytes", encn == LABEL and det(list(raw)) == LABEL)
check("1ª gravação: sem backup e sem .tmp", bk == u"" and prev == u"" and not os.path.exists(p + ".tmp"))
for name, data in cases[1:]:
    open(p, "wb").write(bytes(data))
    bk, encn, prev = w(p, text1, 2)
    check("TXT anterior em %s: backup idêntico byte a byte (sem conversão)" % name,
          open(bk, "rb").read() == bytes(data) and prev == name, (prev, os.path.basename(bk)))
    check("   ... e o novo está em UTF-16 LE com BOM", open(p, "rb").read() == b"\xff\xfe" + text1.encode("utf-16-le"))
obs = os.path.join(d, u"OCA_123_Lista de Materiais_01_obsoletos")
check("backups na pasta '<nome do TXT>_obsoletos' ao lado do TXT", os.path.isdir(obs),
      sorted(os.listdir(d)))
check("nenhuma pasta 'Obsoletos' antiga criada", not os.path.exists(os.path.join(d, "Obsoletos")))
names = sorted(os.listdir(obs))
check("nome do backup = <nome do TXT>_AAAAMMDD_HHMMSS[_n].txt",
      all(__import__("re").match(r"^OCA_123_Lista de Materiais_01_\d{8}_\d{6}(_\d+)?\.txt$", x)
          for x in names), names[:3])
fo, cand = g["obsolete_backup_path"](u"C:/Projetos/OCA_693_Keynote.txt")
check("exemplo do print: OCA_693_Keynote.txt -> OCA_693_Keynote_obsoletos\\OCA_693_Keynote_<data>.txt",
      os.path.basename(fo) == u"OCA_693_Keynote_obsoletos"
      and os.path.basename(cand).startswith(u"OCA_693_Keynote_20"), (fo, cand))
d2 = os.path.join(os.path.dirname(d), "dois"); os.makedirs(d2)
w(os.path.join(d2, "A.txt"), text1, 2); w(os.path.join(d2, "A.txt"), text1, 2)
w(os.path.join(d2, "B.txt"), text1, 2); w(os.path.join(d2, "B.txt"), text1, 2)
check("dois TXT na mesma pasta: cada um com a sua pasta de backup",
      sorted(x for x in os.listdir(d2) if os.path.isdir(os.path.join(d2, x))) == ["A_obsoletos", "B_obsoletos"]
      and len(os.listdir(os.path.join(d2, "A_obsoletos"))) == 1, sorted(os.listdir(d2)))
n_obs = len(os.listdir(obs))
check("nenhum backup sobrescrito (um arquivo por troca)", n_obs == len(cases) - 1, n_obs)

print("\n# 7. Falhas: original sempre mantido")
before = open(p, "rb").read()
for flag, expect in (("copy_backup", u"NÃO foi substituído"), ("corrupt_backup", u"NÃO foi substituído"),
                     ("mkdir", u"NÃO foi substituído"), ("write", u"temporário"),
                     ("locked", u"em uso")):
    shim.FAIL.clear(); shim.FAIL[flag] = True
    if flag == "mkdir":
        shutil.rmtree(obs)
    try:
        w(p, text1.replace(u"PISO", u"PISO NOVO"), 2); msg = None
    except Exception as ex:
        msg = str(ex)
    shim.FAIL.clear()
    check("falha '%s': erro claro, TXT original intacto, sem .tmp" % flag,
          msg is not None and expect in msg and open(p, "rb").read() == before and not os.path.exists(p + ".tmp"), msg)
try:
    w(p, u"A\tB\x07\t\r\n", 1); ok = False
except ValueError:
    ok = True
check("caractere de controle: nada gravado", ok and open(p, "rb").read() == before)
try:
    w(os.path.join(d, "nao_existe", "x.txt"), text1, 2); ok = False
except IOError as ex:
    ok = u"pasta" in str(ex)
check("caminho com pasta inexistente: erro claro", ok)
try:
    w(os.path.join(d, 'Lista<inválida>.txt'), text1, 2); ok = False; m = ""
except IOError as ex:
    ok, m = True, str(ex)
check("nome de arquivo inválido no Windows: erro claro, nada criado",
      ok and not [x for x in os.listdir(d) if "inválida" in x], m)

print("\n# 8. Arquivos pequenos e grandes")
small = u"A\tÚnico\t\r\n"
w(p, small, 1)
check("arquivo de 1 linha", open(p, "rb").read() == b"\xff\xfe" + small.encode("utf-16-le"))
big = u"".join(u"K{0:05d}\tDescrição técnica nº {0} – revestimento cerâmico “A” ≥ {0} mm 😀\t\r\n".format(i)
               for i in range(20000))
t0 = time.time(); w(p, big, 20000); dt = time.time() - t0
raw = open(p, "rb").read()
check("20.000 linhas (%d KB) gravadas, bytes = referência" % (len(raw) // 1024),
      raw == b"\xff\xfe" + big.encode("utf-16-le"), "%.2f s no CPython" % dt)
check("20.000 linhas: Python lê de volta com o codec 'utf-16' (BOM) igual", raw.decode("utf-16") == big)

print("\n# 9. TXT real do escritório (ANSI) convertido para UTF-16 LE com BOM")
if OFFICE:
    o = open(OFFICE, "rb").read()
    check("o TXT do escritório enviado é ANSI (Windows-1252), sem BOM", det(list(o)) == u"ANSI (Windows-1252)", det(list(o)))
    txt = o.decode("cp1252")
    n = txt.count(u"\r\n")
    b, bad = enc(txt)
    check("mesmo texto em UTF-16 LE: bytes = referência, sem caractere bloqueado", b == ref(txt) and not bad)
    check("estrutura do TXT do escritório aprovada (%d linhas)" % n, val(b, n, txt) == [], val(b, n, txt)[:3])
    check("ida e volta: texto idêntico (acentos, cedilha, hierarquia)", dec(b) == txt)

print("\nFunções não encontradas:", missing or "nenhuma")
print("RESULTADO: %d de %d passaram" % (sum(1 for _n, ok in R if ok), len(R)))
