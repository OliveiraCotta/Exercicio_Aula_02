# -*- coding: utf-8 -*-
# ATUALIZAÇÃO DAS DESCRIÇÕES DO MODELO POR KEYNOTE  (pyRevit - Revit 2024/2025/2026)
#
# Lê a planilha de Keynotes do escritório (OCA_XXX_Lista_de_Materiais - aba
# MATERIAIS) e atualiza as descrições dos TIPOS do modelo pelo código de Keynote.
# Fluxo único: EXCEL > MODELO. Escopo: modelo inteiro. Materiais não são lidos.
# (Comentário, não docstring: o tooltip do botão vem SÓ de __doc__ abaixo.)
#
# Onde cada dado mora:
#   * Keynote ........ parâmetro nativo KEYNOTE dos TIPOS usados no modelo.
#   * Descrição PT ... parâmetro nativo "Description" do tipo.
#   * Descrição IN ... parâmetro compartilhado "Descrição IN" do tipo, associado
#                      pelo GUID 70307f6e-3b31-4049-bed4-62dfc3570fec
#                      (OCA_Parametros_Template.txt). OPCIONAL: onde não existe
#                      (ou tem outro GUID) o campo é "Não aplicável" e a
#                      Description continua sendo atualizada. O comando nunca
#                      cria nem vincula esse parâmetro.
#   * Excel .......... aba MATERIAIS: colunas KEYNOTE, DESCRIÇÃO, DESCRIÇÃO IN
#                      (somente leitura - o Excel não é alterado).
#
# Aba 1 - resumo, fluxo (Excel > Modelo), planilha e caminho do TXT de Keynote
#         (sugestão: OCA_XXX_Lista de Materiais_XX.txt, editável).
#
# Aba 2 - janela NÃO modal (dá para navegar no modelo). Três tabelas
# recolhíveis, uma linha por TIPO (família + tipo), colunas ajustáveis, busca
# (botões Buscar / Limpar busca) e "Agrupar por" Sem agrupamento / Keynote /
# Categoria em cada tabela:
#   1. Excel > Modelo: o que o Excel tem e difere do modelo. Valor vazio no
#      Excel nunca apaga o modelo.
#   2. Tipos SEM Keynote ou SEM Descrição: só o Keynote é digitado. A linha
#      fica destacada quando a Keynote existe no Excel (correspondência - não
#      quer dizer gravado). "Verificar no Excel" traz Descrição / Descrição IN
#      da planilha para os campos vazios.
#   3. Tipos com Keynote que não existe no Excel (só consulta).
#   Clicar numa linha seleciona no Revit todas as instâncias do tipo e
#   enquadra na vista ativa (sem transação, a janela continua aberta).
#
# Ao executar a 1ª aba, o TXT de Keynote é gerado do Excel e recarregado no
# projeto em ANSI / Windows-1252 sem BOM (caractere fora do Windows-1252 é
# listado e o TXT NÃO é gravado). O anterior vai para <pasta>\Obsoletos\
# <nome>_AAAAMMDD_HHMMSS.txt, conferido antes da troca; backups nunca são
# apagados nem sobrescritos. Se o Revit recusar o TXT novo, a recarga é desfeita
# e o anterior volta ao lugar (o recusado fica em Obsoletos\..._REJEITADO.txt).
# Filtros por coluna (Agrupamento, Categoria, Família, Tipo, Keynote) e
# cabeçalho fixo: cada tabela rola por dentro.
# Botões da 2ª aba: Voltar (1ª aba) · Dashboard HTML (qualidade dos dados,
# offline) · Atualizar descrições (antes "Gravar marcados": pede confirmação,
# grava e atualiza as tabelas) · Finalizar (só fecha - não grava nem gera nada).
# Visual: Design System OCA (lib/oca_ds) - janelas, diálogos e dashboard.
#
# TXT de Keynote do Revit: mesma estrutura do arquivo do escritório -
# CATEGORIA/prefixos, PREFIXO/GRUPO/CATEGORIA e KEYNOTE / "DESCRIÇÃO |
# DESCRIÇÃO IN" / PREFIXO. Antes de gravar, a árvore é validada (chave vazia
# ou repetida, pai inexistente, ciclos); só é "carregado" depois que o Revit
# confirma a transação e a tabela do projeto bate com o arquivo.
#
# Diagnóstico:
#   * Log por dia em %LOCALAPPDATA%\OCA\pyRevit\logs\Keynotes (cada etapa,
#     transação, exceção com traceback; gravado linha a linha, sobrevive a um
#     encerramento do Revit).
#   * Shift+Clique no botão: modo de diagnóstico com as etapas separadas
#     (limpeza, TXT sem recarga, recarga, análise, janela com dados de teste...).

__title__ = "Atualização\ndas Descrições\npor Keynote"
__persistentengine__ = True        # janela não modal + ExternalEvent
__doc__ = ("Confere as Keynotes dos tipos do modelo com a planilha de materiais "
           "(aba MATERIAIS) e atualiza Description e Descrição IN pelo código de "
           "Keynote.\n\n"
           "Fluxo: Excel > Modelo. Nada é gravado sem confirmação.\n"
           "Inclui conferência em tabelas, seleção no modelo, TXT de Keynote e "
           "dashboard de qualidade.")

import os
import re
import sys
import json
import codecs
from datetime import datetime
from collections import OrderedDict, defaultdict

from pyrevit import revit, DB, script, forms
try:
    from pyrevit import EXEC_PARAMS
    CONFIG_MODE = bool(getattr(EXEC_PARAMS, "config_mode", False))   # Shift+Clique
except Exception:
    CONFIG_MODE = False
try:
    from oca_ds import build_xaml, alert_title, brand_report
except Exception as _ds_ex:
    forms.alert(u"Design System OCA (lib/oca_ds) não encontrado na extensão.",
                sub_msg=u"{}\n\nAtualize a extensão OCA no pyRevit e recarregue.".format(_ds_ex),
                exitscript=True)

import clr
clr.AddReference("System.Xml")
for _asm in ("System.IO.Compression",
             "System.IO.Compression.FileSystem",   # .NET Framework (Revit <= 2024)
             "System.IO.Compression.ZipFile"):     # .NET 8 (Revit 2025+)
    try:
        clr.AddReference(_asm)
    except Exception:
        pass
from System import Guid
from System.IO import StreamReader, File
from System.Text import UTF8Encoding
from System.Xml import XmlDocument, XmlNamespaceManager

doc = revit.doc
BIP = DB.BuiltInParameter
INVALID_ID = DB.ElementId.InvalidElementId
HERE = os.path.dirname(__file__)

output = script.get_output()
output.set_title("Atualização Descrição por Keynotes")
ALERT_TITLE = alert_title(__title__)    # "OCA · Atualização das Descrições por Keynote"


# ------------------------------------------------------------------
# Configuração do escritório
# ------------------------------------------------------------------
SHEET_NAME = u"MATERIAIS"          # aba com os keynotes
DATA_SHEET_NAME = u"DADOS"         # aba PREFIXO -> GRUPO -> CATEGORIA
KN_SHEET_NAME = u"KEYNOTE BÁSICO"  # aba com CATEGORIA -> lista de prefixos (topo do TXT)
NO_GROUP = u"SEM AGRUPAMENTO"
# Parâmetro compartilhado da descrição em inglês (OCA_Parametros_Template.txt).
# A associação é feita pelo GUID - um parâmetro de mesmo nome com outro GUID
# NÃO é considerado (não aparece nas tabelas/etiquetas do template).
PARAM_IN_NAME = u"Descrição IN"
PARAM_IN_GUID_STR = "70307f6e-3b31-4049-bed4-62dfc3570fec"
PARAM_IN_GUID = Guid(PARAM_IN_GUID_STR)
# opções do modo de diagnóstico (Shift+Clique no botão) - padrão = normal
DIAG = {"skip_cleanup": False, "skip_txt_reload": False, "simple_group_rows": False}
NOT_APPLICABLE = u"Não aplicável"   # Descrição IN inexistente no tipo (≠ campo vazio)
MATERIALS_LABEL = u"Materiais"     # categoria usada no relatório p/ materiais

TOOL_TITLE = u"ATUALIZAÇÃO DAS DESCRIÇÕES DO MODELO POR KEYNOTE"
SCRIPT_VERSION = u"2026.10.20"     # aparece na janela e no log: confira se é a cópia nova
# fluxo único: Excel -> Modelo (o Excel é só lido)
MODES = OrderedDict([("to_model", u"Excel → Modelo")])
MODE_DIRS = {"to_model": ("model",)}


# ------------------------------------------------------------------
# Helpers (mesmo padrão das outras ferramentas OCA)
# ------------------------------------------------------------------
# Compatibilidade de motor: o pyRevit roda este script em IronPython 2.7 por
# padrão; se o motor padrão do pyRevit for trocado (IronPython 3.4 / CPython 3),
# 'unicode' e 'unichr' não existem - os nomes abaixo funcionam nos três.
try:
    unicode
except NameError:
    unicode = str
try:
    unichr
except NameError:
    unichr = chr
PY2 = sys.version_info[0] == 2


def to_unicode(value):
    """Coerce any Revit / .NET / byte value into clean unicode text."""
    if value is None:
        return u""
    if isinstance(value, unicode):
        return value
    if isinstance(value, (bytes, bytearray)) and not isinstance(value, unicode):
        for enc in ("utf-8", "cp1252", "latin-1"):
            try:
                return bytes(value).decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
        return bytes(value).decode("ascii", "replace")
    try:
        return unicode(value)
    except Exception:
        return unicode(repr(value))


# ------------------------------------------------------------------
# 0. Log de diagnóstico em arquivo
#   Cada linha é gravada e o arquivo FECHADO na hora (append + close): se o
#   Revit encerrar de forma nativa (sem chance de o Python tratar o erro), a
#   última linha do log mostra a última etapa concluída/iniciada.
#   Pasta: %LOCALAPPDATA%\OCA\pyRevit\logs\Keynotes  (ou a pasta TEMP).
#   Um arquivo por dia; uma execução = um "run" com identificador próprio.
# ------------------------------------------------------------------
import time
import traceback


class RunLog(object):
    def __init__(self):
        self.run_id = datetime.now().strftime("%H%M%S")
        self.path = None
        self.enabled = True
        self._t0 = {}
        for base in (os.environ.get("LOCALAPPDATA"), os.environ.get("APPDATA"),
                     os.environ.get("TEMP"), os.environ.get("TMP")):
            if not base:
                continue
            folder = os.path.join(base, u"OCA", u"pyRevit", u"logs", u"Keynotes")
            try:
                if not os.path.isdir(folder):
                    os.makedirs(folder)
                path = os.path.join(folder, u"keynotes_{}.log".format(
                    datetime.now().strftime("%Y%m%d")))
                with codecs.open(path, "a", encoding="utf-8") as f:
                    f.write(u"")
                self.path = path
                break
            except Exception:
                continue

    def write(self, level, msg):
        """Nunca levanta exceção: se o log falhar, o comando segue."""
        if not self.enabled or not self.path:
            return
        try:
            line = u"{} [{}] {:<5} {}\r\n".format(
                datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3], self.run_id,
                level, to_unicode(msg))
            with codecs.open(self.path, "a", encoding="utf-8") as f:
                f.write(line)
        except Exception:
            pass

    def info(self, msg):
        self.write(u"INFO", msg)

    def warn(self, msg):
        self.write(u"WARN", msg)

    def error(self, msg):
        self.write(u"ERROR", msg)

    def exc(self, where):
        """Exceção Python atual com traceback completo."""
        try:
            tb = to_unicode(traceback.format_exc())
        except Exception:
            tb = u"(traceback indisponível)"
        self.write(u"ERROR", u"{}\r\n{}".format(where, tb.rstrip()))

    def begin(self, step, detail=u""):
        self._t0[step] = time.time()
        self.write(u"STEP", u">> {}{}".format(step, u" - " + detail if detail else u""))

    def end(self, step, detail=u""):
        dt = time.time() - self._t0.pop(step, time.time())
        self.write(u"STEP", u"<< {} ({:.2f} s){}".format(
            step, dt, u" - " + detail if detail else u""))


LOG = RunLog()


def revit_env_info():
    """Versões do Revit (ano + atualização + build), pyRevit, motor Python e
    .NET. Registra no log e devolve um dicionário."""
    info = OrderedDict()
    app = None
    try:
        app = doc.Application
    except Exception:
        pass
    for key, attr in (("revit_version", "VersionNumber"), ("revit_name", "VersionName"),
                      ("revit_build", "VersionBuild"), ("revit_update", "SubVersionNumber")):
        try:
            info[key] = to_unicode(getattr(app, attr))
        except Exception:
            info[key] = u"?"
    # ano principal (2024/2025/2026...) e atualização (ex.: 2025.4)
    major = info.get("revit_version", u"?")
    build = info.get("revit_build", u"")
    m = re.match(r"^(\d{2})\.(\d+)", build or u"")      # builds 24.x / 25.x / 26.x
    if m:
        info["build_family"] = u"{}.x (Revit 20{})".format(m.group(1), m.group(1))
        info["update"] = u"20{}.{}".format(m.group(1), m.group(2))
    else:
        info["build_family"] = u"formato antigo ou desconhecido"
        info["update"] = info.get("revit_update", u"?")
    info["revit_major"] = major
    try:
        from pyrevit import versionmgr
        info["pyrevit"] = to_unicode(versionmgr.get_pyrevit_version().get_formatted())
    except Exception:
        try:
            import pyrevit
            info["pyrevit"] = to_unicode(getattr(pyrevit, "__version__", u"?"))
        except Exception:
            info["pyrevit"] = u"?"
    if info["pyrevit"] in (u"?", u"", u"None"):
        try:                                     # pyRevit 6/7: sessioninfo
            from pyrevit.loader import sessioninfo
            info["pyrevit"] = to_unicode(sessioninfo.get_runtime_info().pyrevit_version)
        except Exception:
            pass
    info["python"] = to_unicode(sys.version.split("\n")[0])
    try:
        import platform
        info["engine"] = u"{} {}".format(platform.python_implementation(),
                                         platform.python_version())
    except Exception:
        info["engine"] = u"IronPython" if sys.platform == "cli" else u"?"
    info["py2"] = to_unicode(PY2)
    try:
        info["persistent_engine"] = to_unicode(bool(__persistentengine__))
    except NameError:
        info["persistent_engine"] = u"(variável ausente)"
    try:
        import System
        info["dotnet"] = to_unicode(System.Environment.Version)
        try:
            from System.Runtime.InteropServices import RuntimeInformation
            info["dotnet_desc"] = to_unicode(RuntimeInformation.FrameworkDescription)
        except Exception:
            pass
        info["os"] = to_unicode(System.Environment.OSVersion)
    except Exception:
        pass
    try:
        info["document"] = to_unicode(doc.Title)
        info["workshared"] = to_unicode(doc.IsWorkshared)
        info["doc_path"] = to_unicode(doc.PathName)
    except Exception:
        pass
    # versões em que este script foi revisado (as demais rodam, mas sem garantia)
    try:
        info["supported"] = u"sim" if int(major) in (2024, 2025, 2026) else u"NÃO testado"
    except Exception:
        info["supported"] = u"?"
    LOG.info(u"Ambiente: " + u" · ".join(u"{}={}".format(k, v) for k, v in info.items()))
    return info


def eid_int(eid):
    """ElementId -> int (Revit 2024+ usa .Value, versões antigas .IntegerValue)."""
    raw = eid.Value if hasattr(eid, "Value") else eid.IntegerValue
    return int(raw)


def clean(text):
    """Texto de parâmetro/célula sem espaços nas pontas. O conteúdo é mantido
    exatamente (todos os dígitos/símbolos do keynote contam)."""
    return to_unicode(text).strip()


_ACCENTS = {u"Á": u"A", u"À": u"A", u"Â": u"A", u"Ã": u"A", u"Ä": u"A",
            u"É": u"E", u"È": u"E", u"Ê": u"E", u"Ë": u"E",
            u"Í": u"I", u"Ì": u"I", u"Î": u"I", u"Ï": u"I",
            u"Ó": u"O", u"Ò": u"O", u"Ô": u"O", u"Õ": u"O", u"Ö": u"O",
            u"Ú": u"U", u"Ù": u"U", u"Û": u"U", u"Ü": u"U", u"Ç": u"C"}


def norm_header(text):
    """'Descrição  IN' -> 'DESCRICAO IN' (maiúsculas, sem acento, 1 espaço)."""
    t = to_unicode(text).upper().replace(u"_", u" ").replace(u"-", u" ")
    t = u"".join(_ACCENTS.get(ch, ch) for ch in t)
    return u" ".join(t.split())


def elem_name(elem):
    try:
        return to_unicode(DB.Element.Name.__get__(elem))
    except Exception:
        try:
            return to_unicode(elem.Name)
        except Exception:
            return u""


def family_name(elem):
    try:
        return to_unicode(elem.FamilyName)
    except Exception:
        return u""


def open_in_browser(path):
    try:
        from System.Diagnostics import Process, ProcessStartInfo
        psi = ProcessStartInfo(path)
        psi.UseShellExecute = True
        Process.Start(psi)
        return True
    except Exception:
        pass
    try:
        from System.Diagnostics import Process
        Process.Start("explorer.exe", '"{}"'.format(path))
        return True
    except Exception:
        pass
    try:
        import webbrowser
        return webbrowser.open("file:///" + path.replace("\\", "/"))
    except Exception:
        return False


# ------------------------------------------------------------------
# 1. Janela de opções (Design System OCA - lib/oca_ds)
# ------------------------------------------------------------------
KN_BODY = u"""
  <StackPanel>
    <HeaderedContentControl Header="RESUMO" Style="{StaticResource oca.Section}">
      <TextBlock Style="{StaticResource oca.Hint}"
                 Text="Lê a aba MATERIAIS da planilha de Keynote e atualiza a Descrição e a Descrição IN dos tipos do modelo pelo código de Keynote. Na próxima etapa você confere as tabelas e grava só o que marcar (com confirmação). Ao executar, o TXT de Keynote é gerado a partir do Excel e recarregado no projeto."/>
    </HeaderedContentControl>

    <HeaderedContentControl Header="FLUXO" Style="{StaticResource oca.Section}">
      <TextBlock Text="Excel  &gt;  Modelo" Style="{StaticResource oca.Value}"/>
    </HeaderedContentControl>

    <HeaderedContentControl Header="PLANILHA EXCEL DE KEYNOTE (.XLSX)" Style="{StaticResource oca.Section}">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="xlsx"/>
        <Button x:Name="browse" Grid.Column="1" Content="Procurar..." Margin="8,0,0,0"/>
      </Grid>
    </HeaderedContentControl>

    <HeaderedContentControl Header="TXT DE KEYNOTE DO REVIT" Style="{StaticResource oca.Section}">
      <StackPanel>
        <TextBlock Style="{StaticResource oca.Hint}" Margin="0,0,0,6"
                   Text="Gerado a partir do Excel e recarregado no projeto ao clicar em Gerar TXT e conferir. Troque XXX (projeto) e XX (revisão) no nome. Se o arquivo já existir, é substituído e o anterior vira um backup com data e hora na mesma pasta."/>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBox x:Name="txt_path"/>
          <Button x:Name="txt_browse" Grid.Column="1" Content="Salvar como..." Margin="8,0,0,0"/>
        </Grid>
      </StackPanel>
    </HeaderedContentControl>
  </StackPanel>"""

KN_FOOTER = u"""
  <Button x:Name="ok" Content="Gerar TXT e conferir" Style="{StaticResource oca.Primary}"/>
  <Button x:Name="cancel" Content="Cancelar" Margin="8,0,0,0"/>"""

KN_XAML = build_xaml(title=__title__, subtitle=__doc__, body=KN_BODY,
                     footer_right=KN_FOOTER, size="M")


TXT_DEFAULT_NAME = u"OCA_XXX_Lista de Materiais_XX.txt"   # XXX = projeto · XX = revisão


def default_txt_path(xlsx_path):
    """Sugestão: OCA_XXX_Lista de Materiais_XX.txt na pasta da planilha
    (o usuário troca XXX / XX no campo antes de executar)."""
    folder = os.path.dirname(xlsx_path) if xlsx_path else u""
    return os.path.join(folder, TXT_DEFAULT_NAME) if folder else TXT_DEFAULT_NAME


def _old_default_txt(xlsx_path):
    """Sugestão antiga (nome da planilha + .txt): não é mais reaproveitada."""
    return os.path.splitext(xlsx_path)[0] + u".txt" if xlsx_path else u""


class KeynoteWindow(forms.WPFWindow):
    def __init__(self, xaml, cfg):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.confirmed = False
        self.result = None
        self.xlsx.Text = cfg["xlsx"]
        saved = cfg["txt_path"]
        if not saved or os.path.normcase(saved) == os.path.normcase(_old_default_txt(cfg["xlsx"])):
            saved = default_txt_path(cfg["xlsx"])
        self.txt_path.Text = saved
        self.browse.Click += self._browse
        self.txt_browse.Click += self._txt_browse
        self.ok.Click += self._ok
        self.cancel.Click += self._cancel

    def _browse(self, sender, args):
        picked = forms.pick_file(file_ext="xlsx",
                                 title="Selecione a planilha de Keynotes")
        if picked:
            cur = to_unicode(self.txt_path.Text).strip()
            self.xlsx.Text = picked
            if not cur or os.path.basename(cur) == TXT_DEFAULT_NAME:
                self.txt_path.Text = default_txt_path(picked)

    def _txt_browse(self, sender, args):
        cur = to_unicode(self.txt_path.Text).strip() or default_txt_path(
            to_unicode(self.xlsx.Text).strip())
        kw = {"file_ext": "txt",
              "default_name": os.path.basename(cur) if cur else u"",
              "init_dir": os.path.dirname(cur) if cur else u""}
        try:
            picked = forms.save_file(title="TXT de Keynotes do Revit", **kw)
        except TypeError:                       # pyRevit antigo sem 'title'
            picked = forms.save_file(**kw)
        if picked:
            self.txt_path.Text = picked

    def _ok(self, sender, args):
        path = to_unicode(self.xlsx.Text).strip().strip('"')
        if not path or not os.path.isfile(path) or not path.lower().endswith(".xlsx"):
            forms.alert(u"Selecione um arquivo .xlsx válido.",
                        sub_msg=u"Use Procurar... para escolher a planilha de Keynote.", title=ALERT_TITLE)
            return
        txt = to_unicode(self.txt_path.Text).strip().strip('"') or default_txt_path(path)
        if not txt.lower().endswith(".txt"):
            txt += u".txt"
        if not os.path.isdir(os.path.dirname(txt) or u"."):
            forms.alert(u"A pasta do TXT de Keynote não existe.", sub_msg=txt, title=ALERT_TITLE)
            return
        self.result = {
            "xlsx": path,
            "mode": "to_model",            # fluxo único
            "scope": "model",              # sempre o modelo inteiro
            "materials": False,            # materiais não são lidos
            "txt_on": True,                # TXT gerado e recarregado ao executar
            "txt_path": txt,
        }
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, args):
        self.confirmed = False
        self.Close()


config = script.get_config()


def ask_options(cfg=None):
    """Mostra a 1ª aba (modal). -> opções escolhidas, ou None se cancelou.
    cfg: valores iniciais; sem cfg usa os da última execução."""
    if cfg is None:
        cfg = {"xlsx": to_unicode(config.get_option("xlsx_path", u"")),
               "txt_path": to_unicode(config.get_option("txt_path", u""))}
    w = KeynoteWindow(KN_XAML, cfg)
    w.ShowDialog()
    if not w.confirmed:
        return None
    config.xlsx_path = w.result["xlsx"]
    config.txt_path = w.result["txt_path"]
    script.save_config()
    return w.result


opts = None                        # opções da 1ª aba (definidas em start_session)
XLSX_PATH = u""
MODE = "to_model"                  # Excel -> Modelo
active_view = doc.ActiveView
use_view = False                   # escopo: modelo inteiro


# ------------------------------------------------------------------
# 2. Leitura do Excel (XML dentro do .xlsx, sem Excel instalado)
# ------------------------------------------------------------------
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def ns_manager(xdoc):
    nsm = XmlNamespaceManager(xdoc.NameTable)
    nsm.AddNamespace("a", NS_MAIN)
    nsm.AddNamespace("r", NS_REL)
    nsm.AddNamespace("p", NS_PKG)
    return nsm


def load_xml(text):
    d = XmlDocument()
    d.PreserveWhitespace = True
    d.LoadXml(text)
    return d


def col_letters(idx):
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA'."""
    idx += 1
    s = u""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = unichr(65 + rem) + s
    return s


def split_ref(ref):
    """'E12' -> (4, 12)."""
    col, num = u"", u""
    for ch in ref:
        if ch.isalpha():
            col += ch
        else:
            num += ch
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1, (int(num) if num else 0)


def zip_read(archive, name):
    entry = archive.GetEntry(name)
    if entry is None:
        return None
    reader = StreamReader(entry.Open(), UTF8Encoding(False))
    try:
        return reader.ReadToEnd()
    finally:
        reader.Dispose()


def resolve_sheets(archive):
    """Return OrderedDict norm(sheet name) -> (name, part path, position)."""
    wb = load_xml(zip_read(archive, "xl/workbook.xml"))
    rels = load_xml(zip_read(archive, "xl/_rels/workbook.xml.rels"))
    nsm_wb, nsm_rel = ns_manager(wb), ns_manager(rels)
    targets = {}
    for rel in rels.SelectNodes("//p:Relationship", nsm_rel):
        tgt = to_unicode(rel.GetAttribute("Target")).replace("\\", "/")
        if tgt.startswith("/"):
            tgt = tgt[1:]
        elif not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        targets[to_unicode(rel.GetAttribute("Id"))] = tgt
    out = OrderedDict()
    for pos, sh in enumerate(wb.SelectNodes("//a:sheets/a:sheet", nsm_wb)):
        name = to_unicode(sh.GetAttribute("name"))
        rid = to_unicode(sh.GetAttribute("id", NS_REL))
        out[norm_header(name)] = (name, targets.get(rid, u""), pos)
    return out


def read_shared_strings(archive):
    text = zip_read(archive, "xl/sharedStrings.xml")
    if not text:
        return []
    d = load_xml(text)
    nsm = ns_manager(d)
    return [u"".join(to_unicode(t.InnerText) for t in si.SelectNodes(".//a:t", nsm))
            for si in d.SelectNodes("//a:si", nsm)]


def parse_sheet(xml_text, shared):
    """-> (grid {row: {col: text}}, rows_with_formula set)."""
    d = load_xml(xml_text)
    nsm = ns_manager(d)
    grid, formula_rows = {}, set()
    for c in d.SelectNodes("//a:sheetData/a:row/a:c", nsm):
        ref = to_unicode(c.GetAttribute("r"))
        if not ref:
            continue
        ci, rn = split_ref(ref)
        if c.SelectSingleNode("a:f", nsm) is not None:
            formula_rows.add(rn)
        ctype = to_unicode(c.GetAttribute("t"))
        val = u""
        if ctype == u"inlineStr":
            val = u"".join(to_unicode(t.InnerText) for t in c.SelectNodes(".//a:t", nsm))
        else:
            v = c.SelectSingleNode("a:v", nsm)
            if v is not None:
                raw = to_unicode(v.InnerText)
                if ctype == u"s":
                    try:
                        val = shared[int(raw)]
                    except Exception:
                        val = u""
                else:
                    val = raw
        if val != u"":
            grid.setdefault(rn, {})[ci] = val
    return grid, formula_rows


HEADER_ALIASES = {
    "key": [u"KEYNOTE", u"KEYNOTES", u"NOTA CHAVE", u"KEY"],
    "pt": [u"DESCRICAO", u"DESCRICAO PT", u"DESCRICAO PORTUGUES", u"DESCRIPTION",
           u"DESCRIPTION PT"],
    "en": [u"DESCRICAO IN", u"DESCRICAO EM IN", u"DESCRICAO EN", u"DESCRICAO INGLES",
           u"DESCRIPTION IN", u"DESCRIPTION EN"],
    "cat": [u"CATEGORIA"],
    "grp": [u"GRUPO"],
    "prefix": [u"PREFIXO"],
    "seq": [u"SEQUENCIA"],
    "obs": [u"OBSERVACOES", u"OBSERVACAO"],
}


def open_xlsx(path):
    """Abre o .xlsx só para leitura, permitindo que o Excel (ou outro programa)
    esteja com o arquivo aberto - ZipFile.OpenRead falha nesse caso. Lê o que
    está SALVO no disco. Erros viram mensagens claras."""
    from System.IO import FileStream, FileMode, FileAccess, FileShare
    from System.IO.Compression import ZipArchive, ZipArchiveMode
    if not os.path.isfile(path):
        raise IOError(u"arquivo não encontrado: {}".format(path))
    if not path.lower().endswith(u".xlsx"):
        raise IOError(u"o arquivo não é .xlsx (salve a planilha como Pasta de Trabalho do "
                      u"Excel .xlsx): {}".format(path))
    share = getattr(FileShare, "ReadWrite") | getattr(FileShare, "Delete")
    try:
        stream = FileStream(path, FileMode.Open, FileAccess.Read, share)
    except Exception as ex:
        raise IOError(u"não foi possível abrir a planilha (sem permissão ou bloqueada): "
                      u"{}".format(to_unicode(ex)))
    try:
        return ZipArchive(stream, ZipArchiveMode.Read, False)
    except Exception as ex:
        try:
            stream.Dispose()
        except Exception:
            pass
        raise IOError(u"a planilha está corrompida ou não é um .xlsx válido: {}".format(
            to_unicode(ex)))


def read_excel(path):
    archive = open_xlsx(path)
    try:
        sheets = resolve_sheets(archive)
        shared = read_shared_strings(archive)

        # aba MATERIAIS; se não existir, primeira aba que tenha coluna KEYNOTE
        candidates = []
        if norm_header(SHEET_NAME) in sheets:
            candidates.append(sheets[norm_header(SHEET_NAME)])
        candidates += [s for k, s in sheets.items() if k != norm_header(SHEET_NAME)]

        chosen = None
        for name, part, pos in candidates:
            text = zip_read(archive, part) if part else None
            if not text:
                continue
            grid, frows = parse_sheet(text, shared)
            for rn in sorted(grid.keys())[:15]:
                heads = dict((ci, norm_header(v)) for ci, v in grid[rn].items())
                if u"KEYNOTE" in heads.values() or u"NOTA CHAVE" in heads.values():
                    chosen = (name, part, pos, grid, frows, rn, heads)
                    break
            if chosen:
                break
        if chosen is None:
            raise Exception(u"Nenhuma aba com a coluna KEYNOTE foi encontrada "
                            u"(procurada primeiro a aba '{}').".format(SHEET_NAME))

        # estrutura de agrupamentos (mesma do TXT de Keynotes do escritório):
        #   DADOS          A=PREFIXO  B=GRUPO  C=CATEGORIA   (linhas "AC  ACESSÓRIOS  ACESSÓRIOS")
        #   KEYNOTE BÁSICO A=CATEGORIA B=prefixos          (linhas "ACESSÓRIOS  AC, DS, ...")
        # valores BRUTOS guardados para o TXT sair idêntico (inclui espaços finais)
        prefix_map, dados_rows, kn_headers = {}, [], []
        dados = sheets.get(norm_header(DATA_SHEET_NAME))
        if dados and dados[1]:
            dtext = zip_read(archive, dados[1])
            if dtext:
                dgrid, _ = parse_sheet(dtext, shared)
                for rn in sorted(dgrid.keys()):
                    cells = dgrid[rn]
                    pf = clean(cells.get(0, u""))
                    if not pf:
                        continue
                    dados_rows.append((to_unicode(cells.get(0, u"")),
                                       to_unicode(cells.get(1, u"")),
                                       to_unicode(cells.get(2, u""))))
                    if pf not in prefix_map:
                        prefix_map[pf] = (clean(cells.get(1, u"")), clean(cells.get(2, u"")))
        kb = sheets.get(norm_header(KN_SHEET_NAME))
        if kb and kb[1]:
            ktext = zip_read(archive, kb[1])
            if ktext:
                kgrid, kfrows = parse_sheet(ktext, shared)
                for rn in sorted(kgrid.keys()):
                    if rn in kfrows:            # a partir daqui são fórmulas =DADOS!...
                        break
                    a_, b_ = kgrid[rn].get(0, u""), kgrid[rn].get(1, u"")
                    if clean(a_) and clean(b_):
                        kn_headers.append((to_unicode(a_), to_unicode(b_)))
        if not kn_headers:                      # sem a aba: monta a partir de DADOS
            order = OrderedDict()
            for pf, grp, cat in dados_rows:
                order.setdefault(cat, []).append(clean(pf))
            kn_headers = [(cat, u", ".join(pfs)) for cat, pfs in order.items()]
    finally:
        archive.Dispose()

    name, part, pos, grid, frows, header_row, heads = chosen
    cols = {}
    for key, aliases in HEADER_ALIASES.items():
        for ci in sorted(heads.keys()):
            if heads[ci] in aliases:
                cols[key] = ci
                break
    missing = [k for k in ("key", "pt", "en") if k not in cols]
    if missing:
        raise Exception(u"Colunas não encontradas na aba '{}': {}".format(
            name, u", ".join({"key": u"KEYNOTE", "pt": u"DESCRIÇÃO",
                              "en": u"DESCRIÇÃO IN"}[m] for m in missing)))

    def cell(cells, key):
        ci = cols.get(key)
        return clean(cells.get(ci, u"")) if ci is not None else u""

    rows, last_row = [], header_row
    for rn in sorted(grid.keys()):
        if rn <= header_row:
            continue
        cells = grid[rn]
        rec = dict((k, cell(cells, k)) for k in HEADER_ALIASES.keys())
        rec["row"] = rn
        # KEYNOTE é fórmula CONCAT(PREFIXO:SEQUENCIA); sem valor em cache
        # (arquivo nunca recalculado) monta a partir das duas colunas.
        if not rec["key"] and rec["prefix"] and rec["seq"]:
            rec["key"] = rec["prefix"] + rec["seq"]
        if any(rec[k] for k in ("key", "pt", "en", "prefix", "seq")):
            last_row = max(last_row, rn)
        if rec["key"]:
            rows.append(rec)

    by_key, dups = OrderedDict(), defaultdict(list)
    for rec in rows:
        if rec["key"] in by_key:
            dups[rec["key"]].append(rec["row"])
        else:
            by_key[rec["key"]] = rec

    return {
        "path": path, "sheet": name, "part": part, "pos": pos,
        "header_row": header_row, "cols": cols, "rows": rows,
        "by_key": by_key, "dups": dict(dups), "last_row": last_row,
        "grid": grid, "formula_rows": frows, "prefix_map": prefix_map,
        "dados_rows": dados_rows, "kn_headers": kn_headers,
        "cat_order": [clean(c) for c, _ in kn_headers],
    }


XL = None                          # planilha lida (start_session)


# ------------------------------------------------------------------
# 3. Keynotes do modelo (tipos usados + materiais usados)
# ------------------------------------------------------------------
def param_text(p):
    if p is None or not p.HasValue:
        return u""
    try:
        return clean(p.AsString())
    except Exception:
        return u""


def keynote_of(elem):
    try:
        return param_text(elem.get_Parameter(BIP.KEYNOTE_PARAM))
    except Exception:
        return u""


def desc_param(elem):
    """Parâmetro nativo Description (independe do idioma do Revit)."""
    try:
        p = elem.get_Parameter(BIP.ALL_MODEL_DESCRIPTION)
        if p is not None:
            return p
    except Exception:
        pass
    for nm in (u"Description", u"Descrição"):
        try:
            p = elem.LookupParameter(nm)
            if p is not None and p.StorageType == DB.StorageType.String:
                return p
        except Exception:
            pass
    return None


def in_param(elem):
    """Parâmetro compartilhado 'Descrição IN', localizado pelo GUID
    (None se o parâmetro não estiver vinculado a este tipo / material)."""
    try:
        p = elem.get_Parameter(PARAM_IN_GUID)
        if p is not None and p.StorageType == DB.StorageType.String:
            return p
    except Exception:
        pass
    return None


def in_param_clash(elem):
    """True se existe um parâmetro chamado 'Descrição IN' que NÃO é o
    compartilhado do escritório (GUID diferente / parâmetro de projeto)."""
    try:
        p = elem.LookupParameter(PARAM_IN_NAME)
        if p is None:
            return False
        return not (p.IsShared and p.GUID == PARAM_IN_GUID)
    except Exception:
        return False


def is_model_category(cat):
    try:
        return cat.CategoryType == DB.CategoryType.Model
    except Exception:
        return False


def collect_model(view_only, include_materials):
    col = (DB.FilteredElementCollector(doc, active_view.Id) if view_only
           else DB.FilteredElementCollector(doc))
    instances = col.WhereElementIsNotElementType()

    type_info, mat_info = {}, {}      # id int -> [ElementId, count, set(categorias)]
    for e in instances:
        try:
            cat = e.Category
            if cat is None:
                continue
            cname = to_unicode(cat.Name)
        except Exception:
            continue
        try:
            tid = e.GetTypeId()
        except Exception:
            tid = None
        if tid is not None and tid != INVALID_ID:
            info = type_info.setdefault(eid_int(tid), [tid, 0, set()])
            info[1] += 1
            info[2].add(cname)
        if include_materials and is_model_category(cat):
            seen = set()
            for paint in (False, True):
                try:
                    mids = list(e.GetMaterialIds(paint))
                except Exception:
                    continue
                for mid in mids:
                    k = eid_int(mid)
                    if k in seen:
                        continue
                    seen.add(k)
                    info = mat_info.setdefault(k, [mid, 0, set()])
                    info[1] += 1
                    info[2].add(cname)

    holders = []

    def add(kind, elem, category, count, used_in):
        key = keynote_of(elem)
        if not key:
            return
        pp, ep = desc_param(elem), in_param(elem)
        holders.append({
            "kind": kind, "elem": elem, "id": eid_int(elem.Id),
            "category": category, "family": family_name(elem) if kind == "type" else u"",
            "name": elem_name(elem), "count": count, "usedIn": sorted(used_in),
            "key": key, "pt_p": pp, "en_p": ep,
            "en_clash": ep is None and in_param_clash(elem),
            "pt": param_text(pp), "en": param_text(ep) if ep is not None else None,
        })

    for tid, count, cats in type_info.values():
        try:
            t = doc.GetElement(tid)
            if t is None:
                continue
            cname = to_unicode(t.Category.Name) if t.Category is not None else sorted(cats)[0]
            add("type", t, cname, count, cats)
        except Exception:
            LOG.exc(u"collect_model: tipo {}".format(eid_int(tid)))
    for mid, count, cats in mat_info.values():
        m = doc.GetElement(mid)
        if m is None:
            continue
        add("material", m, MATERIALS_LABEL, count, cats)

    by_key = OrderedDict()
    for h in sorted(holders, key=lambda x: (x["key"], x["category"], x["name"])):
        by_key.setdefault(h["key"], []).append(h)
    return by_key


KN_PARENT = {}        # chave -> chave pai no arquivo de keynotes carregado no Revit


def keynote_file_texts():
    """Texto (e pai) de cada chave do .txt de keynotes carregado no projeto."""
    out = {}
    try:
        table = DB.KeynoteTable.GetKeynoteTable(doc)
        for entry in table.GetKeyBasedTreeEntries():
            try:
                k = clean(entry.Key)
                out[k] = clean(entry.KeynoteText)
                KN_PARENT[k] = clean(entry.ParentKey)
            except Exception:
                continue
    except Exception:
        pass
    return out


MODEL = OrderedDict()               # lido em start_session
FILE_TEXT = {}


# ------------------------------------------------------------------
# 3b. 'Descrição IN' é OPCIONAL
#   O comando nunca cria nem vincula o parâmetro. Onde ele não existe, o campo
#   é ignorado ("Não aplicável") e a Description continua sendo atualizada.
#   Parâmetro com o MESMO NOME e GUID diferente: só aparece no diagnóstico
#   (log e aviso), nunca recebe gravação.
# ------------------------------------------------------------------
def in_param_stats(scan):
    """-> {"with": n, "without": n, "clash": [categorias]} dos tipos analisados."""
    st = {"with": 0, "without": 0, "clash": set()}
    for info in scan:
        t = info["t"]
        if in_param(t) is not None:
            st["with"] += 1
        else:
            st["without"] += 1
            if in_param_clash(t):
                st["clash"].add(info["category"])
    st["clash"] = sorted(st["clash"])
    return st


# ------------------------------------------------------------------
# 4. Comparação
# ------------------------------------------------------------------
def weighted_values(holders, field):
    """Valores distintos não vazios, do mais usado para o menos usado."""
    w = defaultdict(int)
    for h in holders:
        v = h.get(field)
        if v:
            w[v] += max(1, h["count"])
    return [v for v, _ in sorted(w.items(), key=lambda kv: (-kv[1], kv[0]))]


def split_keynote(key):
    """'RE01' -> ('RE', '01'); 'PCF-01/02' -> ('PCF', '-01/02'). A fórmula
    CONCAT(PREFIXO:SEQUENCIA) do template devolve o keynote original inteiro."""
    i = 0
    while i < len(key) and key[i].isalpha():
        i += 1
    return key[:i], key[i:]


changes = []     # {"target","action","key","field","old","new","where","id"}
skipped = []     # {"key","where","id","eid","reason"}


# ------------------------------------------------------------------
# 5a. Backup de arquivos gravados (TXT de Keynotes)
# ------------------------------------------------------------------
# ------------------------------------------------------------------
# 5b. Excel -> Modelo
# ------------------------------------------------------------------
def owned_by_other(elem):
    if not doc.IsWorkshared:
        return False
    try:
        return (DB.WorksharingUtils.GetCheckoutStatus(doc, elem.Id)
                == DB.CheckoutStatus.OwnedByOtherUser)
    except Exception:
        return False


def where_label(h):
    if h["kind"] == "material":
        return u"Material: {}".format(h["name"])
    fam = h["family"]
    return u"{}: {}{}".format(h["category"], (fam + u" - ") if fam else u"", h["name"])


def plan_to_model():
    """-> (plan, bloqueados). plan: alterações possíveis; bloqueados: diferenças
    que não podem ser gravadas (parâmetro somente leitura, workshared...).
    'Descrição IN' ausente no tipo (ou com GUID diferente) NÃO é bloqueio: o
    campo é simplesmente ignorado e a Description segue normalmente."""
    plan, blocked = [], []
    for key, holders in MODEL.items():
        rec = XL["by_key"].get(key)
        if rec is None:
            continue
        for h in holders:
            for field, label, pkey in (("pt", u"Description", "pt_p"),
                                       ("en", PARAM_IN_NAME, "en_p")):
                new = rec[field]
                if not new:
                    continue                      # Excel vazio nunca apaga o modelo
                p = h[pkey]
                if p is None and field == "en":
                    continue                      # Descrição IN não aplicável neste tipo
                if p is not None:
                    try:
                        label = to_unicode(p.Definition.Name)
                    except Exception:
                        pass
                old = h[field] or u""
                if old == new:
                    continue
                reason = None
                if p is None:
                    reason = (u"parâmetro '{}' não existe neste {}".format(
                        label, u"material" if h["kind"] == "material" else u"tipo"))
                elif p.IsReadOnly:
                    reason = u"parâmetro '{}' somente leitura".format(label)
                elif owned_by_other(h["elem"]):
                    reason = u"elemento emprestado por outro usuário (workshared)"
                item = {"h": h, "field": field, "label": label, "p": p,
                        "old": old, "new": new, "reason": reason}
                (blocked if reason else plan).append(item)
    return plan, blocked


# ------------------------------------------------------------------
# 5c. Agrupamento de Keynotes (estrutura do TXT: CATEGORIA > GRUPO/prefixo)
# ------------------------------------------------------------------
def agrup_of(key):
    """-> (categoria, rótulo do grupo). Ordem de busca: aba DADOS do Excel
    (prefixo), colunas CATEGORIA/GRUPO da linha do keynote, hierarquia do
    arquivo de keynotes carregado no Revit."""
    prefix = split_keynote(key)[0]
    if prefix in XL["prefix_map"]:
        grp, cat = XL["prefix_map"][prefix]
        return (cat or NO_GROUP), (u"{} · {}".format(prefix, grp) if grp else prefix)
    rec = XL["by_key"].get(key)
    if rec is not None and (rec.get("cat") or rec.get("grp")):
        return (rec.get("cat") or NO_GROUP), (
            u"{} · {}".format(prefix, rec["grp"]) if rec.get("grp") else prefix)
    chain, k, seen = [], key, set()
    while KN_PARENT.get(k) and k not in seen:
        seen.add(k)
        k = KN_PARENT[k]
        chain.append(k)
    if chain:
        grp = chain[0]
        label = u"{} · {}".format(grp, FILE_TEXT.get(grp, u"")) if len(chain) > 1 else grp
        return chain[-1], label
    return NO_GROUP, (prefix or u"-")


def cat_rank(cat):
    order = XL["cat_order"]
    if cat in order:
        return order.index(cat)
    return len(order) + (1 if cat == NO_GROUP else 0)


# ------------------------------------------------------------------
# 5d. Situação dos TIPOS no modelo (sem keynote / descrição / descrição IN)
#   O comando NÃO colore mais o modelo: a conferência é feita selecionando a
#   linha da tabela, que seleciona e mostra os elementos na vista ativa.
#   As cores deixadas pelas versões anteriores são localizadas e removidas
#   (seção 9).
# ------------------------------------------------------------------
def view_allows_overrides(view):
    if view is None or getattr(view, "IsTemplate", False):
        return False
    try:
        return bool(view.AreGraphicsOverridesAllowed())
    except Exception:
        pass
    bad = set()
    for _n in ("Schedule", "DrawingSheet", "Report", "ColumnSchedule", "PanelSchedule",
               "Internal", "Undefined", "ProjectBrowser", "SystemBrowser"):
        _vt = getattr(DB.ViewType, _n, None)
        if _vt is not None:
            bad.add(_vt)
    return view.ViewType not in bad


def scan_types(view=None):
    """Tipos COM parâmetro Keynote usados por elementos de categorias de modelo
    (no modelo inteiro ou só os visíveis na vista). -> lista de dicts."""
    try:
        col = (DB.FilteredElementCollector(doc, view.Id) if view is not None
               else DB.FilteredElementCollector(doc))
        elems = col.WhereElementIsNotElementType()
    except Exception:
        LOG.exc(u"scan_types: coletor")
        elems = []
    by_type = OrderedDict()
    n_skip = 0
    for e in elems:
        try:                               # elemento inválido / inacessível: pula
            cat = e.Category
            if cat is None or not is_model_category(cat):
                continue
            tid = e.GetTypeId()
            if tid is None or tid == INVALID_ID:
                continue
            info = by_type.setdefault(eid_int(tid), {"tid": tid, "ids": [],
                                                     "category": to_unicode(cat.Name)})
            info["ids"].append(e.Id)
        except Exception:
            n_skip += 1
    out = []
    for info in by_type.values():
        try:
            t = doc.GetElement(info["tid"])
            if t is None or t.get_Parameter(BIP.KEYNOTE_PARAM) is None:
                continue                   # categoria sem parâmetro Keynote
        except Exception:
            n_skip += 1
            continue
        info["t"] = t
        out.append(info)
    if n_skip:
        LOG.warn(u"scan_types: {} elemento(s)/tipo(s) inacessível(is) ignorado(s)".format(n_skip))
    return out


def type_status(t):
    """-> (keynote, descrição, descrição IN, tem parâmetro IN, campos faltando)"""
    key = keynote_of(t)
    desc = param_text(desc_param(t))
    ip = in_param(t)
    en = param_text(ip) if ip is not None else u""
    checks = [(u"Keynote", key), (u"Descrição", desc)]
    if ip is not None:                    # sem o parâmetro: não aplicável, não é falta
        checks.append((PARAM_IN_NAME, en))
    missing = [lbl for lbl, v in checks if not v]
    return key, desc, en, ip is not None, missing


# ---- cores deixadas no modelo pelas versões anteriores deste comando ----
# Cores usadas antes (laranja, vermelho, lilás, rosa claro, roxo, rosa, azul),
# puras e no tom claro (45% de branco). Os corretos ficavam em meio-tom + 70%.
_OLD_RGB = ((255, 140, 0), (232, 74, 74), (200, 162, 200), (255, 182, 193),
            (142, 68, 173), (255, 105, 180), (52, 152, 219))
OLD_COLORS = set(_OLD_RGB) | set(
    tuple(int(round(c + (255 - c) * 0.45)) for c in rgb) for rgb in _OLD_RGB)
_OGS_COLORS = ("ProjectionLineColor", "CutLineColor", "SurfaceForegroundPatternColor",
               "CutForegroundPatternColor", "SurfaceBackgroundPatternColor",
               "CutBackgroundPatternColor", "ProjectionFillColor", "CutFillColor")


def _rgb(c):
    try:
        if c is not None and c.IsValid:
            return (int(c.Red), int(c.Green), int(c.Blue))
    except Exception:
        pass
    return None


def is_old_highlight(o):
    """Sobreposição igual à que o comando aplicava (2+ campos com a mesma cor
    do comando, ou meio-tom + 70% de transparência)."""
    try:
        if o.Halftone and int(o.Transparency) == 70:
            return True
    except Exception:
        pass
    hits = defaultdict(int)
    for name in _OGS_COLORS:
        try:
            rgb = _rgb(getattr(o, name, None))
        except Exception:
            rgb = None
        if rgb in OLD_COLORS:
            hits[rgb] += 1
    return any(n >= 2 for n in hits.values())


def _colors_log_path(name="keynotes_cores"):
    try:
        return script.get_document_data_file(name, "json")
    except Exception:
        return None


def graphic_views():
    out = []
    try:
        for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
            if view_allows_overrides(v):
                out.append(v)
    except Exception:
        pass
    return out


def find_old_highlight(views, progress=None):
    """-> [(vista, [ElementId com a cor antiga])]"""
    found = []
    for n, v in enumerate(views):
        if progress is not None:
            try:
                progress.update_progress(n + 1, len(views))
            except Exception:
                pass
        ids = []
        try:
            for e in DB.FilteredElementCollector(doc, v.Id).WhereElementIsNotElementType():
                cat = e.Category
                if cat is None or not is_model_category(cat):
                    continue
                try:
                    if is_old_highlight(v.GetElementOverrides(e.Id)):
                        ids.append(e.Id)
                except Exception:
                    continue
        except Exception:
            continue
        if ids:
            found.append((v, ids))
    return found


def clear_old_highlight(found):
    """Volta a sobreposição desses elementos ao padrão. -> (limpos, falhas)"""
    cnt = {"ok": 0, "bad": 0}

    def body():
        for v, ids in found:
            for eid in ids:
                try:
                    v.SetElementOverrides(eid, DB.OverrideGraphicSettings())
                    cnt["ok"] += 1
                except Exception:
                    cnt["bad"] += 1
    ok, why = run_transaction(u"Remover cores antigas - Keynotes", body)
    if not ok:
        LOG.error(u"Limpeza: transação desfeita - " + why)
        return 0, cnt["ok"] + cnt["bad"]
    return cnt["ok"], cnt["bad"]


# ---- ações no Revit com a janela aberta (janela não modal) ----
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent, ExternalEventRequest
from System.Collections.Generic import List


class _RevitAction(IExternalEventHandler):
    """Fila de funções executadas no contexto da API do Revit (ExternalEvent).
    Sem assinatura permanente de eventos do Revit (como Idling): um handler
    que fica assinado pode ser chamado depois de o pyRevit recarregar e
    descartar este motor Python, o que é um risco de encerramento do Revit.
    Antes de cada ação confere se o documento do comando ainda está aberto;
    toda exceção é registrada no log com traceback e mostrada ao usuário -
    nunca sobe para o Revit."""

    def __init__(self):
        self.queue = []
        self.uiapp = None
        self.started = {}           # nome da ação -> momento em que começou

    def Execute(self, uiapp):
        self.run_pending(uiapp, u"ExternalEvent")

    def run_pending(self, uiapp, source):
        self.uiapp = uiapp
        while self.queue:
            name, fn = self.queue.pop(0)
            self.started[name] = time.time()
            LOG.info(u"{}: executando '{}'".format(source, name))
            try:
                if not doc.IsValidObject:
                    LOG.warn(u"{}: documento do comando foi fechado - ação '{}' "
                             u"ignorada".format(source, name))
                    continue
                fn()
                LOG.info(u"{}: '{}' concluída".format(source, name))
            except SystemExit:
                pass
            except Exception as ex:
                LOG.exc(u"{}: '{}' falhou".format(source, name))
                try:
                    forms.alert(u"Não foi possível concluir a ação '{}'.".format(name),
                                sub_msg=u"{}\n\nDetalhes no log:\n{}".format(
                                    to_unicode(ex), LOG.path or u"(log indisponível)"),
                                title=ALERT_TITLE)
                except Exception:
                    pass

    def GetName(self):
        return u"Atualização das Descrições por Keynote"


REVIT_ACTION = _RevitAction()
REVIT_EVENT = ExternalEvent.Create(REVIT_ACTION)
def run_in_revit(fn, name=u"ação"):
    """Enfileira fn para rodar no contexto da API (ExternalEvent).
    -> True se o Revit aceitou o pedido (Accepted/Pending); False se recusou -
    nesse caso a ação sai da fila e nada fica pendente."""
    REVIT_ACTION.queue.append((name, fn))
    REVIT_ACTION.started.pop(name, None)
    res = None
    try:
        res = REVIT_EVENT.Raise()
    except Exception:
        LOG.exc(u"ExternalEvent.Raise('{}')".format(name))
    ok = res in (ExternalEventRequest.Accepted, ExternalEventRequest.Pending)
    (LOG.info if ok else LOG.warn)(u"ExternalEvent.Raise('{}') -> {}".format(name, res))
    if not ok:
        REVIT_ACTION.queue[:] = [q for q in REVIT_ACTION.queue if q[1] is not fn]
    return ok


def action_started(name):
    return name in REVIT_ACTION.started


def _active_uidoc():
    """UIDocument ativo SE for o documento deste comando; senão None."""
    try:
        uiapp = REVIT_ACTION.uiapp
        uidoc = uiapp.ActiveUIDocument if uiapp is not None else revit.uidoc
        if uidoc is None:
            return None
        d = uidoc.Document
        if d.Equals(doc) or (d.PathName == doc.PathName and d.Title == doc.Title):
            return uidoc
    except Exception:
        LOG.exc(u"_active_uidoc")
    return None


def show_in_view(ids):
    """Seleciona os elementos e aproxima na vista ativa (sem transação).
    -> mensagem"""
    if not ids:
        return u""
    uidoc = _active_uidoc()
    if uidoc is None:
        return (u"O documento ativo não é o deste comando ({}). Volte para ele para "
                u"selecionar os elementos.".format(to_unicode(doc.Title)))
    view = uidoc.ActiveView
    # só ids que ainda existem (o usuário pode ter apagado elementos)
    ids = [i for i in ids if doc.GetElement(i) is not None]
    if not ids:
        return u"Os elementos desta linha não existem mais no modelo."
    try:
        uidoc.Selection.SetElementIds(List[DB.ElementId](ids))
    except Exception:
        LOG.exc(u"show_in_view: SetElementIds")
        return u"Não foi possível selecionar os elementos (ver log)."
    try:
        is_graphic = view is not None and not view.IsTemplate and view.ViewType not in (
            DB.ViewType.Schedule, DB.ViewType.DrawingSheet, DB.ViewType.ProjectBrowser,
            DB.ViewType.SystemBrowser, DB.ViewType.Internal, DB.ViewType.Undefined,
            DB.ViewType.Report, DB.ViewType.ColumnSchedule, DB.ViewType.PanelSchedule)
    except Exception:
        is_graphic = False
    if not is_graphic:
        return (u"{} elemento(s) selecionado(s). A vista ativa não é uma vista de modelo; "
                u"abra uma planta, corte ou 3D para enquadrar.".format(len(ids)))
    # só os elementos da linha (não percorre a vista inteira a cada clique)
    try:
        visible = set(eid_int(i) for i in DB.FilteredElementCollector(
            doc, List[DB.ElementId](ids)).WherePasses(
                DB.VisibleInViewFilter(doc, view.Id)).ToElementIds())
    except Exception:
        LOG.exc(u"show_in_view: elementos visíveis")
        visible = set()
    vis = [i for i in ids if eid_int(i) in visible]
    if not vis:
        return (u"{} elemento(s) PRÉ-SELECIONADO(S) no Revit, mas nenhum está visível na vista "
                u"ativa ({}). A seleção continua: abra outra vista (ex.: 3D) - ou clique na "
                u"linha de novo depois de trocar de vista para enquadrar.".format(
                    len(ids), to_unicode(view.Name)))
    try:
        uidoc.ShowElements(List[DB.ElementId](vis))
    except Exception:
        LOG.exc(u"show_in_view: ShowElements")
    try:
        uidoc.Selection.SetElementIds(List[DB.ElementId](ids))   # seleção mantida
    except Exception:
        pass
    return u"{} de {} elemento(s) selecionado(s) e mostrado(s) na vista '{}'.".format(
        len(vis), len(ids), to_unicode(view.Name))


# ------------------------------------------------------------------
# 5e. Tabela 1 - Excel -> Modelo
#     Uma linha por Keynote + Categoria + Família + Tipo: tipos diferentes com
#     o mesmo Keynote NÃO são juntados (as instâncias do tipo ficam na Qtd).
# ------------------------------------------------------------------
def model_rows():
    plan, blocked = plan_to_model()
    by_h = OrderedDict()
    for it in plan + blocked:
        by_h.setdefault(it["h"]["id"], []).append(it)
    merged = OrderedDict()
    for its in by_h.values():
        h = its[0]["h"]
        new = {"pt": None, "en": None}
        blk = {}
        for it in its:
            if it["reason"]:
                blk[it["field"]] = it["reason"]
            else:
                new[it["field"]] = it["new"]
        gkey = (h["key"], h["category"], h["family"], h["name"], h["id"])
        r = merged.get(gkey)
        if r is None:
            r = merged[gkey] = {
                "dir": "model", "key": h["key"], "h": h, "hs": [], "items": [],
                "category": h["category"], "pt_old": h["pt"], "en_old": h["en"],
                "pt_new": new["pt"], "en_new": new["en"], "over": set(), "blocked": blk}
        r["hs"].append(h)
        r["items"].extend(its)
        for it in its:
            if not it["reason"] and it["old"]:
                r["over"].add(it["field"])
    rows = list(merged.values())
    for r in rows:
        r["enabled"] = any(not it["reason"] for it in r["items"])
    return rows, plan, blocked


def _names(values, n=2):
    vals = []
    for v in values:
        if v and v not in vals:
            vals.append(v)
    if len(vals) <= n:
        return u", ".join(vals)
    return u"{} (+{})".format(u", ".join(vals[:n]), len(vals) - n)


def row_view(r):
    """Texto das colunas da tabela 1 (janela e HTML)."""
    cat, grp = agrup_of(r["key"])
    hs = r["hs"]
    fam = _names([h["family"] or h["category"] for h in hs])
    if len(hs) == 1:
        typ = hs[0]["name"]
    else:
        typ = u"{} tipos: {}".format(len(hs), _names([h["name"] for h in hs], 3))

    def cur(v, no_param=False):
        if no_param:
            return NOT_APPLICABLE
        return v or u"(vazio)"

    def new(f):
        if f == "en" and r["h"]["en_p"] is None:
            return NOT_APPLICABLE
        v = r[f + "_new"]
        if v is not None:
            return v or u"(vazio)"
        if f in r["blocked"]:
            return u"(bloqueado)"
        return u"(mantém)"

    dest = u"→ Modelo"
    notes = []
    if r["over"]:
        notes.append(u"sobrescreve")
    for f, why in sorted(r["blocked"].items()):
        notes.append(u"{} bloqueado: {}".format(u"PT" if f == "pt" else u"IN", why))
    if notes:
        dest += u" · " + u" · ".join(notes)
    kind = u"blocked" if not r["enabled"] else (u"over" if r["over"] else u"fill")
    return {"cat": cat, "grp": grp, "category": r["category"], "family": fam,
            "type": typ, "nTypes": len(hs), "count": sum(h["count"] for h in hs),
            "key": r["key"], "ptCur": cur(r["pt_old"]), "ptNew": new("pt"),
            "enCur": cur(r["en_old"], r["h"]["en"] is None), "enNew": new("en"),
            "dest": dest, "dir": r["dir"], "kind": kind}


# ------------------------------------------------------------------
# 5e'. Tabela 2 - tipos do modelo SEM Keynote ou SEM Descrição (editável)
# ------------------------------------------------------------------
def collect_incomplete(t1_keys, scan=None):
    """-> (tabela 2, tabela 3)
    Tabela 2: tipos sem Keynote, ou sem Descrição cujo Keynote está no Excel
              mas NÃO aparece na tabela 1 (o que está na tabela 1 já será
              atualizado por ela).
    Tabela 3: tipos com Keynote preenchido que NÃO existe no Excel (só
              informativo - o código precisa ser incluído na planilha)."""
    inc, notx = [], []
    for info in (scan if scan is not None else scan_types()):
        t = info["t"]
        key, desc, en, has_in, missing = type_status(t)
        row = {"t": t, "id": eid_int(t.Id), "category": info["category"],
               "family": family_name(t), "name": elem_name(t),
               "count": len(info["ids"]), "hasIn": has_in,
               "key0": key, "pt0": desc, "en0": en,
               "key": key, "pt": desc, "en": en,
               "sel": False, "check": u"", "status": u""}
        if key and key not in XL["by_key"]:
            row["check"] = u"Keynote {} não está no Excel".format(key)
            notx.append(row)
            continue
        if key and desc:
            continue
        if key and key in t1_keys:
            continue                         # a tabela 1 já atualiza este keynote
        inc.append(row)
    for lst in (inc, notx):
        lst.sort(key=lambda r: (r["category"], r["family"], r["name"]))
    return inc, notx


def inc_changed(r):
    """Algum campo digitado difere do valor atual? (vazio nunca apaga)"""
    return any(clean(r[f]) and clean(r[f]) != r[f + "0"] for f in ("key", "pt", "en"))


def excel_rec(key):
    """Linha do Excel (aba MATERIAIS) do keynote, ou None (vazio / inexistente)."""
    key = clean(key)
    return XL["by_key"].get(key) if key else None


def verify_incomplete(r, autofill):
    """Confere a linha com o Excel. autofill=True: quando a Keynote existe na
    planilha, PROPÕE a Description e a Descrição IN do Excel (preenche campo
    vazio ou substitui valor diferente), igual à tabela 1. A proposta só é
    gravada se a linha estiver marcada e o usuário confirmar 'Atualizar
    descrições'. Valor vazio no Excel nunca apaga o modelo. Descrição IN só
    entra onde o parâmetro existe."""
    for f in ("key", "pt", "en"):
        r[f] = clean(r[f])
    rec = excel_rec(r["key"])
    applied = []
    if autofill:
        r["pt"], r["en"] = r["pt0"], r["en0"]       # parte sempre do valor do modelo
        if rec is not None:
            if rec["pt"] and rec["pt"] != r["pt0"]:
                r["pt"] = rec["pt"]
                applied.append(u"Description ({})".format(
                    u"substitui" if r["pt0"] else u"preenche"))
            if r["hasIn"] and rec["en"] and rec["en"] != r["en0"]:
                r["en"] = rec["en"]
                applied.append(u"{} ({})".format(
                    PARAM_IN_NAME, u"substitui" if r["en0"] else u"preenche"))
    errs, warns = [], []
    if not r["key"]:
        errs.append(u"falta Keynote")
    elif rec is None:
        warns.append(u"Keynote não está no Excel")
    if not r["pt"]:
        errs.append(u"falta Descrição" + (u" (vazia também no Excel)"
                                          if rec is not None and not rec["pt"] else u""))
    elif rec is not None and rec["pt"] and r["pt"] != rec["pt"]:
        warns.append(u"Descrição difere do Excel")
    if not r["hasIn"]:
        pass                              # Descrição IN não aplicável: nada a avisar
    elif not r["en"]:
        warns.append(u"{} vazia".format(PARAM_IN_NAME))
    elif rec is not None and rec["en"] and r["en"] != rec["en"]:
        warns.append(u"{} difere do Excel".format(PARAM_IN_NAME))
    parts = []
    if rec is not None:
        parts.append(u"Keynote encontrada no Excel (linha {})".format(rec["row"]))
    if applied:
        parts.append(u"Do Excel ao gravar: " + u", ".join(applied))
    if errs:
        parts.append(u"Erro: " + u", ".join(errs))
    if warns:
        parts.append(u"Aviso: " + u", ".join(warns))
    r["check"] = u" · ".join(parts) if parts else u"OK"
    return r["check"]


# ------------------------------------------------------------------
# 5f. Janela de verificação (Aba 2): duas tabelas, busca e agrupamento
# ------------------------------------------------------------------
for _asm in ("PresentationFramework", "PresentationCore", "WindowsBase", "System.Data"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass
import System
from System.Data import DataTable

# cada tabela rola por dentro (o cabeçalho das colunas fica fixo) com altura
# máxima ajustada à janela (_fit_tables); a página rola entre as seções.
# Virtualização ligada: modelos grandes não criam todas as linhas de uma vez.
# (cores, contorno e fonte: estilos implícitos do Design System - lib/oca_ds/wpf.py)
_GRID_STYLE = u"""
              AutoGenerateColumns="False" CanUserAddRows="False"
              CanUserDeleteRows="False" CanUserSortColumns="False" CanUserReorderColumns="False"
              CanUserResizeColumns="True"
              CanUserResizeRows="False" HeadersVisibility="Column" SelectionMode="Single"
              SelectionUnit="FullRow" GridLinesVisibility="Horizontal"
              AlternationCount="2" Margin="0,4,0,0"
              VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
              MinHeight="120" MaxHeight="520"
              EnableRowVirtualization="True" EnableColumnVirtualization="False"
              VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Standard"
              VirtualizingPanel.ScrollUnit="Pixel"
              ScrollViewer.CanContentScroll="True"
              ColumnHeaderStyle="{StaticResource head}" RowStyle="{StaticResource row}" """

# cor da CÉLULA que vai mudar (coluna "<campo>Mark" da linha):
#   fill = preenche vazio · over = sobrescreve · edit = digitado · blocked = não grava
#   excel = vem do Excel e será gravado · diff = Excel difere do modelo (não grava)
# (fundo, texto): pincéis do Design System, que acompanham o tema da janela
MARK_COLORS = (("fill", "{StaticResource oca.OkSoft}", "{StaticResource oca.Ok}"),
               ("over", "{StaticResource oca.WarnSoft}", "{StaticResource oca.Warn}"),
               ("edit", "{StaticResource oca.AccentSoft}", "{StaticResource oca.Violet}"),
               ("excel", "{StaticResource oca.OkSoft}", "{StaticResource oca.Ok}"),
               ("diff", "{StaticResource oca.WarnSoft}", "{StaticResource oca.Warn}"))


def _mark_style(key, mark_col):
    trig = u"".join(
        u'<DataTrigger Binding="{{Binding {0}}}" Value="{1}">'
        u'<Setter Property="Background" Value="{2}"/><Setter Property="Foreground" Value="{3}"/>'
        u'<Setter Property="FontWeight" Value="SemiBold"/></DataTrigger>'.format(mark_col, v, bg, fg)
        for v, bg, fg in MARK_COLORS)
    trig += (u'<DataTrigger Binding="{{Binding {0}}}" Value="blocked">'
             u'<Setter Property="Foreground" Value="{{StaticResource oca.Ink2}}"/></DataTrigger>'.format(mark_col))
    return (u'<Style x:Key="{0}" TargetType="DataGridCell" '
            u'BasedOn="{{StaticResource {{x:Type DataGridCell}}}}">'
            u'<Style.Triggers>{1}</Style.Triggers></Style>'.format(key, trig))


_CHECK_COL = u"""
        <DataGridTemplateColumn Header="✓" Width="36">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <CheckBox HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0"
                        IsChecked="{Binding Sel, UpdateSourceTrigger=PropertyChanged}"
                        IsEnabled="{Binding Enabled}"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>"""


def _col(header, binding, width, style=u"wrap", readonly=True, cell=None):
    """Colunas de texto >= 100 começam proporcionais (ocupam a largura da janela);
    as estreitas (Qtd, Keynote) começam com largura fixa. O usuário pode arrastar
    a divisa do cabeçalho para alargar ou estreitar qualquer coluna."""
    size = (u'Width="{}*" MinWidth="{}"'.format(width, int(width * 0.55)) if width >= 100
            else u'Width="{}"'.format(width))
    return (u'<DataGridTextColumn Header="{}" Binding="{{Binding {}}}" {} '
            u'IsReadOnly="{}" ElementStyle="{{StaticResource {}}}"{}/>'.format(
                header, binding, size, u"True" if readonly else u"False", style,
                u' CellStyle="{{StaticResource {}}}"'.format(cell) if cell else u""))


# estilos referenciados pelas colunas geradas em _col / _GRID_STYLE e pelo corpo
# (vão para Window.Resources, depois dos pincéis e estilos oca.*)
VERIFY_RESOURCES = u"""
    <Style x:Key="wrap" TargetType="TextBlock">
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Padding" Value="4,2,4,2"/>
    </Style>
    <Style x:Key="mono" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="Padding" Value="4,2,4,2"/>
    </Style>
    <Style x:Key="edit" TargetType="TextBlock">
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Padding" Value="4,2,4,2"/>
      <Setter Property="Background" Value="{StaticResource oca.AccentSoft}"/>
      <Setter Property="FontFamily" Value="Consolas"/>
    </Style>
    <Style x:Key="lbl" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource oca.Accent}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,16,0"/>
    </Style>
    <Style x:Key="small" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="sec" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource oca.Surface}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="{StaticResource oca.RadiusGroup}"/>
      <Setter Property="Padding" Value="6,4,8,4"/>
    </Style>
    <Style x:Key="toggle" TargetType="Button" BasedOn="{StaticResource oca.Compact}">
      <Setter Property="Width" Value="28"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
    <Style x:Key="head" TargetType="DataGridColumnHeader" BasedOn="{StaticResource {x:Type DataGridColumnHeader}}"/>
    <Style x:Key="row" TargetType="DataGridRow" BasedOn="{StaticResource {x:Type DataGridRow}}">
      <Setter Property="Background" Value="{StaticResource oca.Field}"/>
      <Style.Triggers>
        <Trigger Property="AlternationIndex" Value="1">
          <Setter Property="Background" Value="{StaticResource oca.Surface2}"/>
        </Trigger>
        <DataTrigger Binding="{Binding Kind}" Value="blocked">
          <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
        </DataTrigger>
        <!-- tabela 2: keynote encontrada no Excel (correspondência, não "gravado") -->
        <DataTrigger Binding="{Binding Kind}" Value="match">
          <Setter Property="Background" Value="{StaticResource oca.AccentSoft}"/>
          <Setter Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
          <Setter Property="BorderThickness" Value="3,0,0,0"/>
        </DataTrigger>
        <!-- linha de título do grupo (Agrupar por), na largura toda da tabela -->
        <DataTrigger Binding="{Binding Kind}" Value="group">
          <Setter Property="Focusable" Value="False"/>
          <Setter Property="IsHitTestVisible" Value="False"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="DataGridRow">
                <Border Background="{StaticResource oca.Surface2}" BorderBrush="{StaticResource oca.Line}"
                        BorderThickness="0,1,0,1" Padding="8,5,8,5">
                  <TextBlock Text="{Binding Grp}" Foreground="{StaticResource oca.Accent}" FontWeight="Bold"/>
                </Border>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </DataTrigger>
        <!-- modo de diagnóstico: título do grupo como linha comum (sem trocar o
             template da linha do DataGrid) -->
        <DataTrigger Binding="{Binding Kind}" Value="grouprow">
          <Setter Property="Focusable" Value="False"/>
          <Setter Property="IsHitTestVisible" Value="False"/>
          <Setter Property="Background" Value="{StaticResource oca.Surface2}"/>
          <Setter Property="Foreground" Value="{StaticResource oca.Accent}"/>
          <Setter Property="FontWeight" Value="Bold"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>
    """ + _mark_style(u"mPt", u"PtMark") + u"""
    """ + _mark_style(u"mEn", u"EnMark") + u"""
    """ + _mark_style(u"mKey", u"KeyMark")

VERIFY_BODY = u"""
  <DockPanel>
    <StackPanel DockPanel.Dock="Top" Margin="0,0,12,0">
      <Border Style="{StaticResource oca.Msg.Info}" Margin="0" Visibility="{Binding Visibility, ElementName=info}">
        <TextBlock x:Name="info" Style="{StaticResource oca.MsgText}"/>
      </Border>
      <WrapPanel Margin="0,10,0,0">
        <TextBlock Text="Buscar" Style="{StaticResource oca.Label}"/>
        <TextBox x:Name="f_text" Width="300" Margin="0,0,8,0"/>
        <Button x:Name="b_search" Content="Buscar" Margin="0,0,8,0"/>
        <Button x:Name="b_clear" Content="Limpar busca" Margin="0,0,8,0"/>
        <TextBlock x:Name="searchinfo" Style="{StaticResource small}" Foreground="{StaticResource oca.Ink2}"
                   Margin="8,0,0,0"/>
      </WrapPanel>
      <Border Style="{StaticResource oca.Msg.Ok}" Visibility="{Binding Visibility, ElementName=viewinfo}">
        <TextBlock x:Name="viewinfo" Style="{StaticResource oca.MsgText}"/>
      </Border>
    </StackPanel>
    <TextBlock x:Name="counter" DockPanel.Dock="Bottom" Margin="0,10,12,0" TextWrapping="Wrap"
               Foreground="{StaticResource oca.Ink}"/>
    <ScrollViewer x:Name="scroll" VerticalScrollBarVisibility="Visible"
                  HorizontalScrollBarVisibility="Disabled" Margin="0,8,0,0">
      <StackPanel Margin="0,0,12,0">
        <Border Style="{StaticResource sec}" Margin="0,4,0,0">
          <DockPanel LastChildFill="False">
            <Button x:Name="tg1" Content="▼" Style="{StaticResource toggle}" ToolTip="Recolher / expandir a tabela"/>
            <TextBlock x:Name="tt1" Style="{StaticResource lbl}" Margin="0,0,8,0" Cursor="Hand"
                       Text="1 · EXCEL &gt; MODELO"/>
            <TextBlock x:Name="h1" Style="{StaticResource small}" Margin="0,0,18,0"
                       Foreground="{StaticResource oca.Ink2}"/>
            <TextBlock Text="Agrupar por " Style="{StaticResource small}"/>
            <ComboBox x:Name="f_mode1" Width="150" Margin="0,0,14,0"/>
          </DockPanel>
        </Border>
        <StackPanel x:Name="body1" Margin="0,6,0,0">
            <TextBlock Style="{StaticResource oca.Hint}"
                       Text="Uma linha por tipo (família + tipo): tipos diferentes com a mesma keynote ficam em linhas separadas. Clique numa linha para selecionar o tipo no Revit.   Célula verde = preenche vazio  ·  Âmbar = substitui o valor atual  ·  Cinza = não pode ser gravado"/>
            <WrapPanel Margin="0,4,0,2"><TextBlock Text="Filtrar:" Style="{StaticResource small}" FontWeight="SemiBold" Margin="0,0,8,0"/><TextBlock Text="Agrupamento" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f1_grp" Width="130" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Categoria" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f1_cat" Width="120" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Família" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f1_fam" Width="130" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Tipo" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f1_typ" Width="140" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Keynote" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f1_key" Width="90" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/></WrapPanel>
            <DataGrid x:Name="grid" """ + _GRID_STYLE + u""">
              <DataGrid.Columns>""" + _CHECK_COL + u"""
                """ + _col(u"Agrupamento", u"Grupo", 150) + u"""
                """ + _col(u"Categoria", u"Categoria", 110) + u"""
                """ + _col(u"Família", u"Familia", 130) + u"""
                """ + _col(u"Tipo", u"Tipo", 170) + u"""
                """ + _col(u"Keynote", u"Keynote", 75, u"mono") + u"""
                """ + _col(u"Descrição atual", u"PtAtual", 180) + u"""
                """ + _col(u"Descrição nova", u"PtNova", 180, cell=u"mPt") + u"""
                """ + _col(u"Descrição em IN atual", u"EnAtual", 160) + u"""
                """ + _col(u"Descrição em IN nova", u"EnNova", 160, cell=u"mEn") + u"""
                """ + _col(u"Destino", u"Destino", 130) + u"""
              </DataGrid.Columns>
            </DataGrid>
        </StackPanel>
        <Border Style="{StaticResource sec}" Margin="0,16,0,0">
          <DockPanel LastChildFill="False">
            <Button x:Name="tg2" Content="▼" Style="{StaticResource toggle}" ToolTip="Recolher / expandir a tabela"/>
            <TextBlock x:Name="tt2" Style="{StaticResource lbl}" Margin="0,0,8,0" Cursor="Hand"
                       Text="2 · ELEMENTOS SEM KEYNOTE OU SEM DESCRIÇÃO"/>
            <TextBlock x:Name="h2" Style="{StaticResource small}" Margin="0,0,18,0"
                       Foreground="{StaticResource oca.Ink2}"/>
            <TextBlock Text="Agrupar por " Style="{StaticResource small}"/>
            <ComboBox x:Name="f_mode2" Width="150" Margin="0,0,14,0"/>
            <Button x:Name="b2_check" Content="Verificar no Excel"/>
          </DockPanel>
        </Border>
        <StackPanel x:Name="body2" Margin="0,6,0,0">
            <TextBlock Style="{StaticResource oca.Hint}"
                       Text="Clique duas vezes no Keynote para digitar. Linha destacada (faixa verde-azulada à esquerda) = keynote encontrada no Excel (só indica correspondência; nada é gravado até 'Atualizar descrições'). Ao digitar a Keynote, a Description e a Descrição em IN do Excel já ficam propostas; marque a linha e use 'Atualizar descrições' (grava Keynote e descrições juntos). 'Verificar no Excel' marca todas as linhas com algo a gravar.   Texto lilás = keynote digitado  ·  Verde = preenche campo vazio  ·  Âmbar = substitui o valor atual"/>
            <WrapPanel Margin="0,4,0,2"><TextBlock Text="Filtrar:" Style="{StaticResource small}" FontWeight="SemiBold" Margin="0,0,8,0"/><TextBlock Text="Agrupamento" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f2_grp" Width="130" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Categoria" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f2_cat" Width="120" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Família" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f2_fam" Width="130" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Tipo" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f2_typ" Width="140" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Keynote" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f2_key" Width="90" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/></WrapPanel>
            <DataGrid x:Name="grid2" """ + _GRID_STYLE + u""">
              <DataGrid.Columns>""" + _CHECK_COL + u"""
                """ + _col(u"Categoria", u"Categoria", 105) + u"""
                """ + _col(u"Família", u"Familia", 125) + u"""
                """ + _col(u"Tipo", u"Tipo", 160) + u"""
                """ + _col(u"Qtd", u"Qtd", 45, u"mono") + u"""
                """ + _col(u"Keynote ✎", u"Keynote", 95, u"edit", False, u"mKey") + u"""
                """ + _col(u"Descrição atual", u"PtAtual", 160) + u"""
                """ + _col(u"Descrição Excel", u"PtExcel", 170, cell=u"mPt") + u"""
                """ + _col(u"Descrição em IN atual", u"EnAtual", 140) + u"""
                """ + _col(u"Descrição em IN Excel", u"EnExcel", 150, cell=u"mEn") + u"""
                """ + _col(u"Verificação", u"Verificacao", 230) + u"""
              </DataGrid.Columns>
            </DataGrid>
        </StackPanel>
        <Border Style="{StaticResource sec}" Margin="0,16,0,0">
          <DockPanel LastChildFill="False">
            <Button x:Name="tg3" Content="▼" Style="{StaticResource toggle}" ToolTip="Recolher / expandir a tabela"/>
            <TextBlock x:Name="tt3" Style="{StaticResource lbl}" Margin="0,0,8,0" Cursor="Hand"
                       Text="3 · ELEMENTOS COM KEYNOTE QUE NÃO ESTÁ NO EXCEL"/>
            <TextBlock x:Name="h3" Style="{StaticResource small}" Margin="0,0,18,0"
                       Foreground="{StaticResource oca.Ink2}"/>
            <TextBlock Text="Agrupar por " Style="{StaticResource small}"/>
            <ComboBox x:Name="f_mode3" Width="150" Margin="0,0,14,0"/>
          </DockPanel>
        </Border>
        <StackPanel x:Name="body3" Margin="0,6,0,8">
            <TextBlock Style="{StaticResource oca.Hint}"
                       Text="Só consulta: inclua o código na planilha do Excel e rode o comando de novo."/>
            <WrapPanel Margin="0,4,0,2"><TextBlock Text="Filtrar:" Style="{StaticResource small}" FontWeight="SemiBold" Margin="0,0,8,0"/><TextBlock Text="Agrupamento" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f3_grp" Width="130" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Categoria" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f3_cat" Width="120" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Família" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f3_fam" Width="130" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Tipo" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f3_typ" Width="140" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/><TextBlock Text="Keynote" Style="{StaticResource small}" Margin="0,0,4,0"/><TextBox x:Name="f3_key" Width="90" Margin="0,0,10,0" ToolTip="Filtra a coluna (parte do texto, sem acento). Enter aplica."/></WrapPanel>
            <DataGrid x:Name="grid3" """ + _GRID_STYLE + u""" IsReadOnly="True">
              <DataGrid.Columns>
                """ + _col(u"Categoria", u"Categoria", 120) + u"""
                """ + _col(u"Família", u"Familia", 150) + u"""
                """ + _col(u"Tipo", u"Tipo", 220) + u"""
                """ + _col(u"Qtd", u"Qtd", 45, u"mono") + u"""
                """ + _col(u"Keynote", u"Keynote", 95, u"mono") + u"""
                """ + _col(u"Descrição atual", u"Descricao", 240) + u"""
                """ + _col(u"Descrição em IN atual", u"DescricaoIN", 220) + u"""
                """ + _col(u"Situação", u"Verificacao", 240) + u"""
              </DataGrid.Columns>
            </DataGrid>
        </StackPanel>
      </StackPanel>
    </ScrollViewer>
  </DockPanel>"""

# rodapé: Voltar à esquerda; à direita o botão principal primeiro
VERIFY_FOOTER_LEFT = u"""
  <Button x:Name="b_back" Content="←  Voltar" Style="{StaticResource oca.Ghost}"
          ToolTip="Volta para a primeira aba (planilha e TXT)"/>"""

VERIFY_FOOTER = u"""
  <Button x:Name="b_apply" Content="Atualizar descrições" MinWidth="150" Style="{StaticResource oca.Primary}"
          ToolTip="Grava no modelo as linhas marcadas (pede confirmação)"/>
  <Button x:Name="b_dash" Content="Dashboard HTML" MinWidth="130" Margin="8,0,0,0"
          ToolTip="Abre no navegador o painel de qualidade dos dados da última análise (não altera o modelo)"/>
  <Button x:Name="b_finish" Content="Finalizar" MinWidth="110" Margin="8,0,0,0"
          ToolTip="Fecha o comando sem gravar nada"/>"""

_VERIFY_KW = dict(title=__title__ + u" · Verificação",
                  subtitle=u"Excel > Modelo: confira as tabelas e grave só o que marcar "
                           u"(nada é gravado sem confirmação).  ·  versão " + SCRIPT_VERSION,
                  body=VERIFY_BODY, footer_right=VERIFY_FOOTER,
                  footer_left=VERIFY_FOOTER_LEFT, resources=VERIFY_RESOURCES,
                  size="L", height=980, width=1450)
try:                                       # mínimos da janela (oca_ds mais novo)
    VERIFY_XAML = build_xaml(min_width=900, min_height=620, **_VERIFY_KW)
except TypeError:                          # oca_ds mais antigo, sem min_width
    VERIFY_XAML = build_xaml(**_VERIFY_KW)

GROUP_MODES = [u"Sem agrupamento", u"Keynote", u"Categoria"]
COLS1 = ("Grupo", "Categoria", "Familia", "Tipo", "Keynote", "PtAtual", "PtNova",
         "EnAtual", "EnNova", "Destino")
COLS2 = ("Categoria", "Familia", "Tipo", "Qtd", "Keynote", "PtAtual", "PtExcel",
         "EnAtual", "EnExcel", "Verificacao")
COLS3 = ("Categoria", "Familia", "Tipo", "Qtd", "Keynote", "Descricao", "DescricaoIN",
         "Verificacao")
MARKS = ("PtMark", "EnMark", "KeyMark")
FILTER_CODES = ("grp", "cat", "fam", "typ", "key")   # filtros por coluna (f<n>_<código>)

NO_KEY = u"(sem keynote)"
TYPE_INST = {}        # id do tipo -> [ElementId das instâncias] (modelo inteiro)


def _fold(text):
    """Texto para busca: minúsculo e sem acento ('Descrição' acha 'descricao')."""
    return norm_header(text).lower()


def _key_label(key):
    if not key:
        return NO_KEY
    cat, grp = agrup_of(key)
    if cat == NO_GROUP:                   # prefixo sem agrupamento na aba DADOS
        return key
    return u"{}   ·   {} ({})".format(key, grp, cat)


def _ask(owner, text, title=ALERT_TITLE):
    """Pergunta Sim/Não (janela WPF; pode ser chamada com a tabela aberta)."""
    from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
    try:
        res = MessageBox.Show(owner, text, title, MessageBoxButton.YesNo,
                              MessageBoxImage.Question)
    except Exception:
        res = MessageBox.Show(text, title, MessageBoxButton.YesNo, MessageBoxImage.Question)
    return res == MessageBoxResult.Yes


def _tell(owner, text, title=ALERT_TITLE, warn=False):
    from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
    img = MessageBoxImage.Warning if warn else MessageBoxImage.Information
    try:
        MessageBox.Show(owner, text, title, MessageBoxButton.OK, img)
    except Exception:
        MessageBox.Show(text, title, MessageBoxButton.OK, img)


class VerifyWindow(forms.WPFWindow):
    """Janela NÃO modal: o Revit continua utilizável. Tudo que usa a API do
    Revit (selecionar/mostrar, gravar, carregar o TXT) passa por run_in_revit
    (ExternalEvent). Busca e agrupamento são feitos aqui, em Python: cada
    tabela é remontada com as linhas que atendem à busca e, se houver
    agrupamento, com uma linha de título por grupo (Kind = "group")."""

    def __init__(self, rows, inc, notx, info_text, demo=False):
        LOG.begin(u"Interface", u"montando a janela de verificação")
        forms.WPFWindow.__init__(self, VERIFY_XAML, literal_string=True)
        self.closing_for_back = False
        self.demo = demo                  # dados de teste: nunca grava
        self._busy = False
        self._h_changed1 = self._safe(self._on_changed1, u"Tabela 1: marcar linha")
        self._h_changed2 = self._safe(self._on_changed2, u"Tabela 2: editar linha")
        self.rows, self.inc, self.notx = rows, inc, notx
        self.by_rid = dict((r["rid"], r) for r in rows)
        self.by_rid2 = dict((r["rid"], r) for r in inc)
        self.by_rid3 = dict((r["rid"], r) for r in notx)
        self.info.Text = info_text
        self.viewinfo.Text = u""
        self.searchinfo.Text = u""
        self._query = u""                 # termo aplicado pelo botão Buscar
        self._colf = {1: {}, 2: {}, 3: {}}  # filtros por coluna aplicados
        self._last_ids = []
        self.t1 = self.t2 = self.t3 = None
        self.d1, self.d2, self.d3 = {}, {}, {}
        self._vis = {1: [], 2: [], 3: []}  # RowId das linhas visíveis (sem títulos)
        for combo, idx in ((self.f_mode1, 1), (self.f_mode2, 2), (self.f_mode3, 1)):
            combo.ItemsSource = GROUP_MODES
            combo.SelectedIndex = idx      # tab. 1 e 3: Keynote · tab. 2: Categoria
        self._rebuild()

        # eventos (ligados UMA vez; as tabelas novas ligam só o ColumnChanged).
        # Todos passam por _safe: exceção num evento da janela NÃO MODAL subiria
        # para o Revit e encerraria o programa.
        sf = self._safe
        self.f_mode1.SelectionChanged += sf(self._on_mode, u"Agrupar por (tabela 1)")
        self.f_mode2.SelectionChanged += sf(self._on_mode, u"Agrupar por (tabela 2)")
        self.f_mode3.SelectionChanged += sf(self._on_mode, u"Agrupar por (tabela 3)")
        self.b_search.Click += sf(self._search, u"Buscar")
        self.b_clear.Click += sf(self._clear_search, u"Limpar busca")
        self.f_text.KeyDown += sf(self._on_key, u"Busca (tecla)")
        for n in (1, 2, 3):
            for code in FILTER_CODES:
                box = getattr(self, "f%d_%s" % (n, code), None)
                if box is not None:
                    box.KeyDown += sf(self._on_key, u"Filtro (tecla)")
        self.Loaded += sf(self._fit_tables, u"Ajustar tabelas")
        self.scroll.SizeChanged += sf(self._fit_tables, u"Ajustar tabelas")
        self.grid.SelectionChanged += sf(self._on_select1, u"Seleção tabela 1")
        self.grid2.SelectionChanged += sf(self._on_select2, u"Seleção tabela 2")
        self.grid3.SelectionChanged += sf(self._on_select3, u"Seleção tabela 3")
        # recolher / expandir: botão ▼ ou clique no título da tabela
        for n in (1, 2, 3):
            tog = sf((lambda k: lambda s, a: self._toggle(k))(n), u"Recolher/expandir")
            getattr(self, "tg%d" % n).Click += tog
            getattr(self, "tt%d" % n).MouseLeftButtonUp += tog
        # a roda do mouse sobre uma tabela rola a janela (as tabelas não rolam)
        for g in (self.grid, self.grid2, self.grid3):
            g.PreviewMouseWheel += sf(self._wheel, u"Rolagem")
        self.b2_check.Click += sf(self._verify_all, u"Verificar no Excel")
        self.b_back.Click += sf(self._back, u"Voltar")
        self.b_dash.Click += sf(self._dashboard, u"Dashboard HTML")
        self.b_apply.Click += sf(self._apply, u"Atualizar descrições")
        self.b_finish.Click += sf(self._finish, u"Finalizar")
        self.Closed += sf(self._on_closed, u"Fechar janela")
        self._watch_unhandled(True)
        LOG.end(u"Interface", u"tabela 1: {} · tabela 2: {} · tabela 3: {} linha(s){}".format(
            len(rows), len(inc), len(notx), u" (DADOS DE TESTE)" if demo else u""))

    # ---- proteção da interface ----
    def _safe(self, fn, name):
        """Envolve um evento da interface: registra e mostra o erro em vez de
        deixar a exceção chegar ao Revit."""
        win = self

        def handler(sender, args):
            try:
                return fn(sender, args)
            except Exception as ex:
                LOG.exc(u"Interface: '{}'".format(name))
                try:
                    win.viewinfo.Text = u"Erro em '{}': {} (detalhes no log)".format(
                        name, to_unicode(ex))
                except Exception:
                    pass
        return handler

    def _watch_unhandled(self, on):
        """Registra no log exceções da interface que escapam do Python (ex.:
        dentro do DataGrid do WPF). Se a janela deste comando está em uso,
        marca como tratada para o Revit não ser encerrado."""
        try:
            import System
            from System.Windows.Threading import DispatcherUnhandledExceptionEventHandler
            if on:
                self._disp_h = DispatcherUnhandledExceptionEventHandler(self._on_ui_exception)
                self.Dispatcher.UnhandledException += self._disp_h
                self._dom_h = System.UnhandledExceptionEventHandler(self._on_domain_exception)
                System.AppDomain.CurrentDomain.UnhandledException += self._dom_h
            else:
                if getattr(self, "_disp_h", None) is not None:
                    self.Dispatcher.UnhandledException -= self._disp_h
                    self._disp_h = None
                if getattr(self, "_dom_h", None) is not None:
                    System.AppDomain.CurrentDomain.UnhandledException -= self._dom_h
                    self._dom_h = None
        except Exception:
            LOG.exc(u"_watch_unhandled({})".format(on))

    def _on_ui_exception(self, sender, e):
        try:
            LOG.error(u"Exceção não tratada na interface WPF:\r\n{}".format(
                to_unicode(e.Exception)))
            if self.IsVisible and (self.IsActive or self.IsMouseOver or
                                   self.IsKeyboardFocusWithin):
                e.Handled = True
                self.viewinfo.Text = (u"Erro interno da interface (registrado no log, o Revit "
                                      u"continua aberto): {}".format(to_unicode(e.Exception.Message)))
        except Exception:
            pass

    @staticmethod
    def _on_domain_exception(sender, e):
        try:
            LOG.error(u"Exceção não tratada (AppDomain, terminando={}):\r\n{}".format(
                e.IsTerminating, to_unicode(e.ExceptionObject)))
        except Exception:
            pass

    def _on_closed(self, sender, args):
        LOG.info(u"Janela de verificação fechada{}".format(
            u" (Voltar)" if self.closing_for_back else u""))
        self._watch_unhandled(False)

    # ---- montagem das tabelas ----
    @staticmethod
    def _new_table(cols):
        tb = DataTable("t")
        for name, typ in (("RowId", System.Int32), ("Sel", System.Boolean),
                          ("Enabled", System.Boolean), ("Kind", System.String),
                          ("Grp", System.String)) + tuple((c, System.String)
                                                         for c in cols + MARKS):
            tb.Columns.Add(name, clr.GetClrType(typ))
        return tb

    @staticmethod
    def _mode(combo):
        return GROUP_MODES[min(max(combo.SelectedIndex, 0), len(GROUP_MODES) - 1)]

    @staticmethod
    def _grp_of(mode, key, category):
        if mode == u"Keynote":
            return _key_label(clean(key))
        if mode == u"Categoria":
            return category or u"(sem categoria)"
        return u""

    def _matches(self, n, vals, grp, key):
        """Busca geral (todas as colunas) E filtros por coluna da tabela n."""
        if self._query:
            hay = u" ".join(to_unicode(v) for k, v in vals.items() if not k.startswith(u"_"))
            if self._query not in _fold(hay + u" " + grp):
                return False
        flt = self._colf.get(n) or {}
        if not flt:
            return True
        fields = {"grp": _key_label(clean(key)) + u" " + to_unicode(vals.get("Grupo", u"")),
                  "cat": vals.get("Categoria", u""), "fam": vals.get("Familia", u""),
                  "typ": vals.get("Tipo", u""), "key": vals.get("Keynote", u"")}
        for code, q in flt.items():
            if q not in _fold(fields.get(code, u"")):
                return False
        return True

    def _read_filters(self):
        """Lê as caixas de filtro das três tabelas -> {n: {código: texto}}."""
        out = {}
        for n in (1, 2, 3):
            d = {}
            for code in FILTER_CODES:
                box = getattr(self, "f%d_%s" % (n, code), None)
                if box is None:
                    continue
                v = _fold(to_unicode(box.Text).strip())
                if v:
                    d[code] = v
            out[n] = d
        return out

    def _fill(self, n, tb, store, items, mode, key_of, values_of, sort_key):
        """Monta a tabela n: filtra pela busca, ordena e, se houver
        agrupamento, insere a linha de título de cada grupo. Cada registro
        aparece uma única vez."""
        store.clear()
        entries = []
        for r in items:
            g = self._grp_of(mode, key_of(r), r["category"])
            vals = values_of(r)
            if self._matches(n, vals, g, key_of(r)):
                entries.append((g, r, vals))
        if mode == GROUP_MODES[0]:
            entries.sort(key=lambda e: sort_key(e[1]))
        else:
            entries.sort(key=lambda e: (e[0] == NO_KEY, e[0]) + sort_key(e[1]))
        counts = defaultdict(int)
        for g, _r, _v in entries:
            counts[g] += 1
        last, gi = None, 0
        vis = []
        for g, r, vals in entries:
            if mode != GROUP_MODES[0] and g != last:
                last, gi = g, gi + 1
                hrow = tb.NewRow()
                label = u"{}      ·  {} linha(s)".format(g, counts[g])
                hrow["RowId"], hrow["Sel"], hrow["Enabled"] = -gi, False, False
                hrow["Grp"] = label
                if DIAG["simple_group_rows"]:
                    hrow["Kind"] = u"grouprow"
                    hrow[(COLS1, COLS2, COLS3)[n - 1][0]] = label
                else:
                    hrow["Kind"] = u"group"
                tb.Rows.Add(hrow)
            row = tb.NewRow()
            row["RowId"], row["Grp"] = r["rid"], g
            vals = dict(vals)
            row["Kind"] = vals.pop("_kind")
            row["Sel"], row["Enabled"] = bool(r["sel"]), bool(r.get("enabled", True))
            for c, v in vals.items():
                row[c] = v
            tb.Rows.Add(row)
            store[r["rid"]] = row
            vis.append(r["rid"])
        self._vis[n] = vis

    @staticmethod
    def _mark1(r, f):
        if r[f + "_new"] is not None:
            return u"over" if f in r["over"] else u"fill"
        return u"blocked" if f in r["blocked"] else u""

    def _vals1(self, r):
        v = r["view"]
        d = dict(zip(COLS1, (v["grp"], v["category"], v["family"], v["type"], v["key"],
                             v["ptCur"], v["ptNew"], v["enCur"], v["enNew"], v["dest"])),
                 _kind=v["kind"])
        d.update(PtMark=self._mark1(r, "pt"), EnMark=self._mark1(r, "en"), KeyMark=u"")
        return d

    @staticmethod
    def _cmp2(r, f, rec, has_param=True):
        """-> (texto da coluna 'Excel', marca da célula) para o campo f."""
        if not has_param:
            return NOT_APPLICABLE, u""
        if rec is None:
            return (u"(keynote não está no Excel)" if clean(r["key"]) else u""), u""
        xl = rec[f]
        if r[f] and r[f] != r[f + "0"]:
            # vem do Excel e será gravado: verde preenche vazio, âmbar substitui
            return r[f], (u"over" if r[f + "0"] else u"excel")
        if not xl:
            return u"(vazio no Excel)", u""
        if r[f + "0"] and xl != r[f + "0"]:
            return xl, u"diff"                 # Excel difere do modelo (mantido)
        return xl, u""

    def _vals2(self, r):
        rec = excel_rec(r["key"])
        pt_x, pt_m = self._cmp2(r, "pt", rec)
        en_x, en_m = self._cmp2(r, "en", rec, r["hasIn"])
        d = dict(zip(COLS2, (r["category"], r["family"], r["name"], to_unicode(r["count"]),
                             r["key"], r["pt0"] or u"(vazio)", pt_x,
                             (r["en0"] or u"(vazio)") if r["hasIn"] else NOT_APPLICABLE,
                             en_x, r["check"])),
                 _kind=u"match" if rec is not None else
                 (u"changed" if inc_changed(r) else u"inc"))
        d.update(KeyMark=u"edit" if r["key"] and r["key"] != r["key0"] else u"",
                 PtMark=pt_m, EnMark=en_m)
        return d

    def _vals3(self, r):
        d = dict(zip(COLS3, (r["category"], r["family"], r["name"], to_unicode(r["count"]),
                             r["key"], r["pt"] or u"(vazio)",
                             (r["en"] or u"(vazio)") if r["hasIn"] else NOT_APPLICABLE,
                             r["check"])), _kind=u"info")
        d.update(PtMark=u"", EnMark=u"", KeyMark=u"")
        return d

    @staticmethod
    def _bind(grid, tb):
        """Troca a fonte da grade por uma tabela NOVA (nunca apaga e refaz as
        linhas da tabela ligada, o que deixava a grade vazia)."""
        grid.ItemsSource = None
        try:
            grid.Items.GroupDescriptions.Clear()
        except Exception:
            pass
        grid.ItemsSource = tb.DefaultView

    def _rebuild(self):
        """Remonta as 3 tabelas com o agrupamento e a busca atuais."""
        self._commit()
        keep = dict((n, self._selected_rid(g)) for n, g in
                    ((1, self.grid), (2, self.grid2), (3, self.grid3)))
        self._busy = True
        try:
            self.t1 = self._new_table(COLS1)
            self.t2 = self._new_table(COLS2)
            self.t3 = self._new_table(COLS3)
            self._fill(1, self.t1, self.d1, self.rows, self._mode(self.f_mode1),
                       lambda r: r["key"], self._vals1,
                       lambda r: (r["key"], r["view"]["category"], r["view"]["family"],
                                  r["view"]["type"]))
            self._fill(2, self.t2, self.d2, self.inc, self._mode(self.f_mode2),
                       lambda r: r["key"], self._vals2,
                       lambda r: (r["category"], r["family"], r["name"]))
            self._fill(3, self.t3, self.d3, self.notx, self._mode(self.f_mode3),
                       lambda r: r["key"], self._vals3,
                       lambda r: (r["key"], r["category"], r["family"], r["name"]))
            for n, grid, tb in ((1, self.grid, self.t1), (2, self.grid2, self.t2),
                                (3, self.grid3, self.t3)):
                self._bind(grid, tb)
                self._reselect(grid, keep[n])
            self.t1.ColumnChanged += self._h_changed1
            self.t2.ColumnChanged += self._h_changed2
        except Exception as ex:
            LOG.exc(u"Montar tabelas")
            self.viewinfo.Text = u"Erro ao montar as tabelas: {}".format(to_unicode(ex))
        finally:
            self._busy = False
        self._update_counter()

    @staticmethod
    def _reselect(grid, rid):
        """Mantém destacada a linha que estava selecionada (sem reenviar ao Revit)."""
        if rid is None:
            return
        try:
            for drv in grid.ItemsSource:
                if int(drv.Row["RowId"]) == rid:
                    grid.SelectedItem = drv
                    return
        except Exception:
            pass

    def _on_mode(self, sender, args):
        self._rebuild()

    # ---- busca: só ao clicar em Buscar (ou Enter); Limpar busca restaura ----
    def _search(self, sender, args):
        """Aplica a busca geral e os filtros por coluna (Buscar ou Enter)."""
        self._query = _fold(to_unicode(self.f_text.Text).strip())
        self._colf = self._read_filters()
        self._rebuild()
        n_f = sum(len(v) for v in self._colf.values())
        if self._query or n_f:
            self.searchinfo.Text = u"{}{}{} linha(s) encontrada(s)".format(
                u"Busca '{}' · ".format(to_unicode(self.f_text.Text).strip())
                if self._query else u"",
                u"{} filtro(s) de coluna · ".format(n_f) if n_f else u"",
                sum(len(v) for v in self._vis.values()))
        else:
            self.searchinfo.Text = u""

    def _clear_search(self, sender, args):
        """Limpa a busca geral e todos os filtros por coluna."""
        self.f_text.Text = u""
        for n in (1, 2, 3):
            for code in FILTER_CODES:
                box = getattr(self, "f%d_%s" % (n, code), None)
                if box is not None:
                    box.Text = u""
        self._query = u""
        self._colf = {1: {}, 2: {}, 3: {}}
        self.searchinfo.Text = u""
        self._rebuild()

    def _on_key(self, sender, args):
        try:
            from System.Windows.Input import Key
            if args.Key in (Key.Enter, Key.Return):
                args.Handled = True
                self._search(None, None)
        except Exception:
            pass

    # ---- recolher / expandir e rolagem da janela ----
    def _toggle(self, n):
        from System.Windows import Visibility
        body, btn = getattr(self, "body%d" % n), getattr(self, "tg%d" % n)
        show = body.Visibility != Visibility.Visible
        body.Visibility = Visibility.Visible if show else Visibility.Collapsed
        btn.Content = u"▼" if show else u"▶"

    @staticmethod
    def _inner_scroll(grid):
        """ScrollViewer interno do DataGrid (onde ficam as linhas)."""
        try:
            from System.Windows.Media import VisualTreeHelper
            from System.Windows.Controls import ScrollViewer
            stack = [grid]
            while stack:
                el = stack.pop(0)
                if isinstance(el, ScrollViewer):
                    return el
                for i in range(VisualTreeHelper.GetChildrenCount(el)):
                    stack.append(VisualTreeHelper.GetChild(el, i))
        except Exception:
            pass
        return None

    def _wheel(self, sender, args):
        """Roda do mouse sobre uma tabela: rola as linhas da tabela; quando ela
        chega ao início/fim, rola a página (sem barras brigando)."""
        sv = self._inner_scroll(sender)
        if sv is not None and sv.ScrollableHeight > 0:
            up = args.Delta > 0
            if (up and sv.VerticalOffset > 0) or (not up and sv.VerticalOffset <
                                                  sv.ScrollableHeight):
                return                     # a própria tabela rola
        args.Handled = True
        self.scroll.ScrollToVerticalOffset(self.scroll.VerticalOffset - args.Delta * 0.4)

    def _fit_tables(self, sender=None, args=None):
        """Altura máxima de cada tabela = área visível da página menos os
        títulos: uma tabela aberta ocupa a tela com o cabeçalho sempre à vista."""
        try:
            h = self.scroll.ViewportHeight
            if not h or h != h:            # 0 ou NaN antes de a janela aparecer
                return
            mh = max(220.0, h - 150.0)
            for g in (self.grid, self.grid2, self.grid3):
                g.MaxHeight = mh
        except Exception:
            LOG.exc(u"_fit_tables")

    def _commit(self):
        """Confirma células/linhas em edição (tabelas 1 e 2) antes de ler os
        valores ou remontar as tabelas."""
        try:
            from System.Windows.Controls import DataGridEditingUnit
        except Exception:
            return
        for g in (self.grid, self.grid2):
            try:
                g.CommitEdit(DataGridEditingUnit.Cell, True)
                g.CommitEdit(DataGridEditingUnit.Row, True)
            except Exception:
                try:
                    g.CancelEdit()
                except Exception:
                    pass

    # ---- Revit: selecionar e mostrar na vista ativa (via ExternalEvent) ----
    @staticmethod
    def _ids_of_types(type_ids):
        out = []
        for tid in type_ids:
            out.extend(TYPE_INST.get(tid, []))
        return out

    @staticmethod
    def _selected_rid(grid):
        """RowId da linha selecionada; None para nada / linha de título."""
        item = grid.SelectedItem
        if item is None:
            return None
        try:
            rid = int(item["RowId"])
        except Exception:
            return None
        return rid if rid >= 0 else None

    def _on_select1(self, sender, args):
        rid = None if self._busy else self._selected_rid(self.grid)
        if rid is not None and rid in self.by_rid:
            self._show(self._ids_of_types([h["id"] for h in self.by_rid[rid]["hs"]]))

    def _on_select2(self, sender, args):
        rid = None if self._busy else self._selected_rid(self.grid2)
        if rid is not None and rid in self.by_rid2:
            self._show(self._ids_of_types([self.by_rid2[rid]["id"]]))

    def _on_select3(self, sender, args):
        rid = None if self._busy else self._selected_rid(self.grid3)
        if rid is not None and rid in self.by_rid3:
            self._show(self._ids_of_types([self.by_rid3[rid]["id"]]))

    def _show(self, ids):
        """Seleciona no Revit TODAS as instâncias do tipo da linha (substitui a
        seleção anterior) e enquadra na vista ativa. Sem transação."""
        self._last_ids = list(ids)
        if not ids:
            self.viewinfo.Text = u"Nenhuma instância deste tipo no modelo."
            return
        win = self

        def act():
            win.viewinfo.Text = show_in_view(win._last_ids)
        if not run_in_revit(act, u"Selecionar na vista"):
            self.viewinfo.Text = (u"O Revit não aceitou o pedido agora (comando ou edição em "
                                  u"andamento?). Termine a operação no Revit e clique de novo.")

    # ---- tabela 1 ----
    def _check1(self, rid, value):
        r = self.by_rid[rid]
        r["sel"] = bool(value) and r["enabled"]
        if rid in self.d1:
            self.d1[rid]["Sel"] = r["sel"]

    def _on_changed1(self, sender, e):
        if self._busy or e.Column.ColumnName != "Sel":
            return
        rid = int(e.Row["RowId"])
        if rid < 0:
            return
        self._busy = True
        try:
            self._check1(rid, e.Row["Sel"] == True)
            self._update_counter()
        finally:
            self._busy = False

    # ---- tabela 2 (só o Keynote é digitado; Descrições vêm do Excel) ----
    def _check2(self, rid, value):
        r = self.by_rid2[rid]
        r["sel"] = bool(value)
        if rid in self.d2:
            self.d2[rid]["Sel"] = r["sel"]

    def _refresh2(self, r):
        row = self.d2.get(r["rid"])
        if row is None:                   # fora da busca atual
            return
        vals = self._vals2(r)
        row["Kind"] = vals.pop("_kind")
        for k, v in vals.items():
            row[k] = v

    def _on_changed2(self, sender, e):
        name = e.Column.ColumnName
        if self._busy or name not in ("Sel", "Keynote"):
            return
        rid = int(e.Row["RowId"])
        if rid < 0:
            return
        self._busy = True
        try:
            if name == "Sel":
                self._check2(rid, e.Row["Sel"] == True)
            else:
                r = self.by_rid2[rid]
                r["key"] = clean(e.Row["Keynote"] or u"")
                # novo keynote: confere no Excel e já propõe Description / IN da
                # planilha (gravados só ao confirmar 'Atualizar descrições')
                verify_incomplete(r, autofill=True)
                self._refresh2(r)
                self._check2(rid, inc_changed(r))
            self._update_counter()
        finally:
            self._busy = False

    def _verify_all(self, sender, args):
        """Confere de novo cada Keynote no Excel, propõe Description /
        Descrição IN da planilha e MARCA as linhas que têm algo a gravar.
        Gravar é outro passo (Atualizar descrições + confirmação)."""
        self._commit()
        self._busy = True
        n_match = 0
        try:
            for r in self.inc:
                verify_incomplete(r, autofill=True)
                if excel_rec(r["key"]) is not None:
                    n_match += 1
                self._refresh2(r)
                if inc_changed(r):
                    self._check2(r["rid"], True)
            self._update_counter()
        finally:
            self._busy = False
        self.viewinfo.Text = (u"Verificado no Excel: {} de {} linha(s) com keynote encontrada. "
                              u"Nada foi gravado - confira e use 'Atualizar descrições'.".format(
                                  n_match, len(self.inc)))

    # ---- comum ----
    def _update_counter(self):
        s1, s2 = self.selected(), self.selected_inc()
        tot1 = len([r for r in self.rows if r["enabled"]])
        n_over = len([r for r in s1 if r["over"]])
        n_edit = len([r for r in s2 if inc_changed(r)])
        self.counter.Text = (
            u"Tabela 1: {} de {} linha(s) marcada(s) ({} sobrescreve(m))   ·   "
            u"Tabela 2: {} de {} marcada(s), {} com alteração   ·   "
            u"Tabela 3: {} tipo(s) com Keynote fora do Excel".format(
                len(s1), tot1, n_over, len(s2), len(self.inc), n_edit, len(self.notx)))
        tots = (len(self.rows), len(self.inc), len(self.notx))
        for n, lbl, extra in ((1, self.h1, u"{} marcada(s)".format(len(s1))),
                              (2, self.h2, u"{} marcada(s)".format(len(s2))),
                              (3, self.h3, u"")):
            vis, tot = len(self._vis[n]), tots[n - 1]
            txt = u"{} linha(s)".format(tot) if vis == tot else \
                u"{} de {} linha(s) na busca".format(vis, tot)
            lbl.Text = u"·  " + txt + (u"  ·  " + extra if extra else u"")

    def selected(self):
        return [r for r in self.rows if r["sel"] and r["enabled"]]

    def selected_inc(self):
        return [r for r in self.inc if r["sel"]]

    def reload(self, rows, inc, notx, info_text, msg):
        """Depois de gravar: novas linhas (o que ainda falta) nas 3 tabelas.
        Busca e agrupamento escolhidos continuam valendo."""
        self.rows, self.inc, self.notx = rows, inc, notx
        self.by_rid = dict((r["rid"], r) for r in rows)
        self.by_rid2 = dict((r["rid"], r) for r in inc)
        self.by_rid3 = dict((r["rid"], r) for r in notx)
        self.info.Text = info_text
        self.viewinfo.Text = msg
        self._rebuild()
        self.b_apply.IsEnabled = True

    def _apply(self, sender, args):
        """Grava as linhas marcadas - SÓ depois da confirmação - e atualiza as
        tabelas (a janela continua aberta)."""
        self._commit()
        if self.demo:
            _tell(self, u"Modo de teste com dados fictícios: nada é gravado no modelo.")
            return
        sel1, sel2 = self.selected(), [r for r in self.selected_inc() if inc_changed(r)]
        if not sel1 and not sel2:
            n_marked2 = len(self.selected_inc())
            _tell(self, u"Nada para gravar.\n\n" + (
                u"As {} linha(s) marcada(s) da tabela 2 não têm valor novo: digite a Keynote "
                u"e/ou use 'Verificar no Excel' (o Excel precisa ter a descrição)."
                .format(n_marked2) if n_marked2 else
                u"Marque linhas na tabela 1 ou na tabela 2."))
            self.viewinfo.Text = u"Nada marcado para gravar."
            return
        n_over = len([r for r in sel1 if r["over"]])
        cnt = pending_fields(sel1, sel2)
        if not _ask(self, u"Gravar no modelo?\n\n"
                          u"Tabela 1 (Excel > Modelo): {} tipo(s){}\n"
                          u"Tabela 2 (Keynote / descrições): {} tipo(s)\n\n"
                          u"Campos que serão gravados:\n"
                          u"  • Description: {}\n  • {}: {}\n  • Keynote: {}\n\n"
                          u"Ctrl+Z no Revit desfaz a gravação.".format(
                              len(sel1), u" - {} substitui(em) valor atual".format(n_over)
                              if n_over else u"", len(sel2), cnt["pt"], PARAM_IN_NAME,
                              cnt["en"], cnt["key"])):
            self.viewinfo.Text = u"Gravação cancelada - nada foi alterado."
            return
        self.b_apply.IsEnabled = False
        self.viewinfo.Text = u"Gravando no modelo..."
        LOG.info(u"Atualizar descrições confirmado: tabela 1 {} · tabela 2 {}".format(
            len(sel1), len(sel2)))
        if not run_in_revit(lambda: save_run(sel1, sel2), u"Atualizar descrições"):
            self.b_apply.IsEnabled = True
            self.viewinfo.Text = (u"O Revit não aceitou a gravação agora (comando ou edição em "
                                  u"andamento?). Nada foi gravado - tente de novo.")
            return
        self._watch_action(u"Atualizar descrições", 8)

    def _watch_action(self, name, seconds):
        """Se a ação não começar em alguns segundos, avisa na janela e no log
        (o Revit só executa ações externas quando está livre)."""
        try:
            from System import TimeSpan
            from System.Windows.Threading import DispatcherTimer
        except Exception:
            return
        timer = DispatcherTimer()
        timer.Interval = TimeSpan.FromSeconds(seconds)
        win = self

        def tick(sender, args):
            timer.Stop()
            try:
                if action_started(name):
                    return
                LOG.warn(u"'{}' ainda não começou após {} s (ExternalEvent pendente)".format(
                    name, seconds))
                win.viewinfo.Text = (
                    u"O Revit ainda não executou '{}'. Ele só executa quando está livre: "
                    u"termine o comando ativo (Esc), feche diálogos abertos e passe o mouse "
                    u"sobre a área de desenho. A gravação acontece assim que o Revit "
                    u"liberar.".format(name))
            except Exception:
                LOG.exc(u"_watch_action")
        timer.Tick += tick
        timer.Start()

    def _dashboard(self, sender, args):
        """Dashboard HTML com os dados da última análise (não lê nem grava o
        modelo de novo)."""
        try:
            path = write_dashboard()
        except Exception as ex:
            LOG.exc(u"Dashboard")
            _tell(self, u"Não foi possível gerar o dashboard.\n\n{}".format(to_unicode(ex)),
                  warn=True)
            return
        if open_in_browser(path):
            self.viewinfo.Text = u"Dashboard aberto no navegador: {}".format(path)
        else:
            self.viewinfo.Text = u"Dashboard gerado (abra no navegador): {}".format(path)

    def _back(self, sender, args):
        """Volta para a 1ª aba. O que já foi gravado continua no modelo;
        marcações não gravadas são descartadas."""
        self._commit()
        pend = len(self.selected()) + len([r for r in self.selected_inc() if inc_changed(r)])
        if pend and not _ask(self, u"Voltar para a primeira aba?\n\n{} linha(s) marcada(s) "
                                   u"ainda não foram gravadas e serão descartadas.".format(pend)):
            return
        self.closing_for_back = True
        self.Close()
        run_in_revit(go_back, u"Voltar")

    def _finish(self, sender, args):
        """Encerra o comando: só fecha a interface. Não grava, não relê o
        modelo, não gera TXT nem relatório."""
        LOG.info(u"Finalizar")
        self.Close()
        try:
            output.close()
        except Exception:
            pass


# ------------------------------------------------------------------
# 5g. TXT de Keynotes do Revit a partir do Excel
#     CATEGORIA <tab> lista de prefixos <tab>          (aba KEYNOTE BÁSICO)
#     PREFIXO   <tab> GRUPO <tab> CATEGORIA            (aba DADOS)
#     KEYNOTE   <tab> DESCRIÇÃO | DESCRIÇÃO IN <tab> PREFIXO   (aba MATERIAIS)
# ------------------------------------------------------------------
def _txt_field(text):
    return u" ".join(to_unicode(text).replace(u"\t", u" ").split())


TXT_BUILD = {"entries": []}       # última montagem: [(chave, texto, pai, origem)]


def build_keynote_txt():
    """-> (texto, avisos). Estrutura idêntica ao TXT do escritório. As linhas
    passam por validate_keynote_entries antes de virar texto; cada entrada
    guarda a origem (aba/linha do Excel) para as mensagens de erro."""
    entries, warns = [], []
    for cat, plist in XL["kn_headers"]:
        entries.append((cat, plist, u"", u"aba {}".format(KN_SHEET_NAME)))
    known = set()
    for pf, grp, cat in XL["dados_rows"]:
        entries.append((pf, grp, cat, u"aba {}".format(DATA_SHEET_NAME)))
        known.add(clean(pf))
    seen = set(known) | set(clean(c) for c, _ in XL["kn_headers"])
    for rec in XL["rows"]:
        key = rec["key"]
        if key in seen:
            if key not in known:
                warns.append(u"{}: repetido no Excel (só a 1ª linha entra no TXT)".format(key))
            continue
        seen.add(key)
        text = u" | ".join(_txt_field(v) for v in (rec["pt"], rec["en"]) if _txt_field(v))
        prefix = split_keynote(key)[0]
        parent = prefix if prefix in known else u""
        if not parent:
            warns.append(u"{}: prefixo '{}' não está em DADOS - entra sem agrupamento".format(
                key, prefix))
        entries.append((_txt_field(key), text, parent,
                        u"aba {} linha {}".format(XL.get("sheet", SHEET_NAME), rec["row"])))
    entries, vwarns = validate_keynote_entries(entries)
    warns.extend(vwarns)
    TXT_BUILD["entries"] = entries
    # texto como está no Excel (só tabulação/quebra de linha viram espaço), para
    # o arquivo continuar idêntico ao TXT do escritório
    lines = [u"{}\t{}\t{}".format(*(_no_tab(v) for v in e[:3])) for e in entries]
    return u"\r\n".join(lines) + u"\r\n", warns


# ---- Windows-1252 (ANSI) sem depender do codec do motor Python ----
# 0x00-0x7F e 0xA0-0xFF são iguais ao Unicode; 0x80-0x9F têm 27 caracteres
# próprios (€ “ ” – — … etc.). Os 5 códigos indefinidos (81, 8D, 8F, 90, 9D)
# e qualquer outro caractere NÃO são representáveis: nada vira "?".
_CP1252_HIGH = {0x20AC: 0x80, 0x201A: 0x82, 0x0192: 0x83, 0x201E: 0x84, 0x2026: 0x85,
                0x2020: 0x86, 0x2021: 0x87, 0x02C6: 0x88, 0x2030: 0x89, 0x0160: 0x8A,
                0x2039: 0x8B, 0x0152: 0x8C, 0x017D: 0x8E, 0x2018: 0x91, 0x2019: 0x92,
                0x201C: 0x93, 0x201D: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97,
                0x02DC: 0x98, 0x2122: 0x99, 0x0161: 0x9A, 0x203A: 0x9B, 0x0153: 0x9C,
                0x017E: 0x9E, 0x0178: 0x9F}
_CP1252_BACK = dict((v, k) for k, v in _CP1252_HIGH.items())
# sugestões para os casos mais comuns em descrições técnicas
_CP1252_HINT = {0x2265: u">=", 0x2264: u"<=", 0x2248: u"~", 0x2260: u"<>", 0x03BC: u"µ",
                0x2212: u"-", 0x00A0: u" ", 0x2032: u"'", 0x2033: u'"', 0x2715: u"x",
                0x2716: u"x", 0x2192: u"->", 0x2190: u"<-", 0x221E: u"infinito",
                0x2205: u"Ø", 0x2300: u"Ø", 0x03A9: u"Ohm", 0x2126: u"Ohm"}


def cp1252_encode(text):
    """-> (lista de bytes, [(índice, caractere)] não representáveis)."""
    out, bad = [], []
    for i, ch in enumerate(text):
        o = ord(ch)
        if o in (9, 10, 13) or 0x20 <= o < 0x7F or 0xA0 <= o <= 0xFF:
            out.append(o)
        elif o == 0x7F or o < 0x20:
            bad.append((i, ch))           # caractere de controle (ex.: vindo do Excel)
        elif o in _CP1252_HIGH:
            out.append(_CP1252_HIGH[o])
        else:
            bad.append((i, ch))
    return out, bad


def cp1252_decode(data):
    return u"".join(unichr(_CP1252_BACK.get(b, b)) for b in data)


def non_cp1252_report(entries):
    """Caracteres que impedem gravar o TXT em Windows-1252.
    -> [{"char", "code", "key", "origin", "hint"}]"""
    found = []
    for k, t, p, origin in [(e + (u"",))[:4] for e in entries]:
        for field in (k, t, p):
            _b, bad = cp1252_encode(_no_tab(field))
            for _i, ch in bad:
                found.append({"char": ch, "code": u"U+{:04X}".format(ord(ch)), "key": _txt_field(k),
                              "origin": origin,
                              "hint": _CP1252_HINT.get(ord(ch), u"")})
    return found


def _bytes_of(path):
    return [int(b) for b in File.ReadAllBytes(path)]


def validate_txt_bytes(data, expected_lines):
    """Confere o arquivo gravado: Windows-1252 sem BOM, quebra de linha CRLF,
    3 campos por linha (chave, texto, pai) e a quantidade de entradas.
    -> lista de problemas (vazia = ok)."""
    probs = []
    if data[:3] == [0xEF, 0xBB, 0xBF] or data[:2] in ([0xFF, 0xFE], [0xFE, 0xFF]):
        probs.append(u"o arquivo começa com BOM (deveria ser ANSI puro)")
    for b in (0x81, 0x8D, 0x8F, 0x90, 0x9D):
        if b in data:
            probs.append(u"byte 0x{:02X} não existe no Windows-1252".format(b))
    text = cp1252_decode(data)
    if u"\r\n" not in text and text:
        probs.append(u"quebras de linha não são CRLF")
    body = text.replace(u"\r\n", u"\n")
    if u"\r" in body:
        probs.append(u"há CR solto (quebra de linha inválida)")
    lines = body.split(u"\n")
    if lines and lines[-1] == u"":
        lines = lines[:-1]
    if len(lines) != expected_lines:
        probs.append(u"{} linha(s) no arquivo, esperado {}".format(len(lines), expected_lines))
    for n, ln in enumerate(lines, 1):
        if not ln.strip():
            probs.append(u"linha {} vazia".format(n))
        elif ln.count(u"\t") != 2:
            probs.append(u"linha {} com {} tabulação(ões) (esperado 2)".format(n, ln.count(u"\t")))
        if len(probs) > 20:
            break
    return probs


OBSOLETE_DIR = u"Obsoletos"        # subpasta dos backups, ao lado do TXT


def obsolete_backup_path(path):
    """<pasta do TXT>\\Obsoletos\\<nome>_AAAAMMDD_HHMMSS.txt - nunca repete um nome
    existente (acrescenta _2, _3...)."""
    folder = os.path.join(os.path.dirname(path), OBSOLETE_DIR)
    base, ext = os.path.splitext(os.path.basename(path))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    cand, n = os.path.join(folder, u"{}_{}{}".format(base, stamp, ext)), 1
    while File.Exists(cand):
        n += 1
        cand = os.path.join(folder, u"{}_{}_{}{}".format(base, stamp, n, ext))
    return folder, cand


def _io_message(ex):
    """Texto claro para os erros de arquivo mais comuns."""
    name = type(ex).__name__
    txt = to_unicode(ex)
    if u"UnauthorizedAccess" in name or u"denied" in txt.lower() or u"negado" in txt.lower():
        return u"sem permissão para gravar nessa pasta ou arquivo somente leitura ({})".format(txt)
    if u"being used" in txt or u"sendo usado" in txt or u"IOException" in name:
        return u"arquivo em uso ou bloqueado por outro programa ({})".format(txt)
    if u"DirectoryNotFound" in name or u"could not find" in txt.lower():
        return u"pasta não encontrada ({})".format(txt)
    return txt


def write_keynote_txt(path, text, expected_lines):
    """Grava o TXT em ANSI / Windows-1252 SEM BOM, com segurança:
      1. codifica (para se houver caractere fora do Windows-1252);
      2. grava um .tmp na mesma pasta e confere os bytes e a estrutura;
      3. se o TXT já existe: copia para <pasta>\\Obsoletos\\<nome>_<data_hora>.txt
         e confere a cópia - se o backup falhar, o TXT original NÃO é trocado;
      4. substitui o TXT e confere os bytes finais (se falhar, restaura o backup).
    -> (backup, codificação). Exceções com mensagem clara."""
    import System
    data, bad = cp1252_encode(text)
    if bad:
        raise ValueError(u"{} caractere(s) fora do Windows-1252 - o TXT não foi gravado".format(
            len(bad)))
    folder = os.path.dirname(path) or u"."
    if not System.IO.Directory.Exists(folder):
        raise IOError(u"pasta do TXT não encontrada: {}".format(folder))
    arr = System.Array[System.Byte](data)
    tmp = path + u".tmp"
    try:
        File.WriteAllBytes(tmp, arr)
    except Exception as ex:
        raise IOError(u"não foi possível gravar o arquivo temporário: " + _io_message(ex))
    try:
        if _bytes_of(tmp) != data:
            raise IOError(u"conferência do arquivo temporário falhou (bytes diferentes)")
        probs = validate_txt_bytes(data, expected_lines)
        if probs:
            raise IOError(u"estrutura inválida: " + u"; ".join(probs[:5]))
        backup = u""
        if File.Exists(path):
            obs, backup = obsolete_backup_path(path)
            try:
                System.IO.Directory.CreateDirectory(obs)
                original = _bytes_of(path)
                File.Copy(path, backup, False)          # nunca sobrescreve um backup
                if _bytes_of(backup) != original:
                    raise IOError(u"a cópia de backup não confere com o original")
            except Exception as ex:
                LOG.exc(u"TXT: backup")
                try:                    # cópia ruim não pode parecer backup válido
                    if File.Exists(backup):
                        bad_name = os.path.splitext(backup)[0] + u"_FALHOU.txt"
                        File.Move(backup, bad_name)
                        LOG.warn(u"TXT: cópia de backup que não conferiu renomeada: " + bad_name)
                except Exception:
                    LOG.exc(u"TXT: renomear backup que falhou")
                raise IOError(u"backup em '{}' falhou - o TXT original NÃO foi substituído: "
                              u"{}".format(obs, _io_message(ex)))
            LOG.info(u"TXT: backup conferido em {}".format(backup))
        try:
            File.Copy(tmp, path, True)
        except Exception as ex:
            raise IOError(u"não foi possível substituir o TXT (o original continua): " +
                          _io_message(ex))
        if _bytes_of(path) != data:
            if backup:
                try:
                    File.Copy(backup, path, True)
                    LOG.warn(u"TXT: gravação não conferiu - original restaurado do backup")
                except Exception:
                    LOG.exc(u"TXT: restaurar backup")
            raise IOError(u"o TXT gravado não confere com o gerado")
    finally:
        try:
            if File.Exists(tmp):
                File.Delete(tmp)
        except Exception:
            LOG.exc(u"TXT: apagar temporário")
    LOG.info(u"TXT: gravado e conferido {} ({} bytes, Windows-1252 sem BOM, {} linhas)".format(
        path, len(data), expected_lines))
    return backup, u"ANSI (Windows-1252)"


def _no_tab(text):
    return to_unicode(text).replace(u"\t", u" ").replace(u"\r", u" ").replace(u"\n", u" ")


def validate_keynote_entries(entries):
    """Garante uma árvore de keynotes válida para o Revit: sem chave vazia,
    sem chave repetida, pai existente e sem ciclos (A > B > A ou A > A).
    Linha com problema é corrigida (vai para o nível de cima) ou descartada,
    sempre com aviso. -> (entradas, avisos)"""
    warns, out, keys = [], [], set()
    nk = _txt_field                        # comparação sem espaços extras
    for ent in entries:
        k, t, p = ent[0], ent[1], ent[2]
        origin = ent[3] if len(ent) > 3 else u""
        if not nk(k):
            warns.append(u"linha sem keynote descartada ('{}')".format(nk(t)[:30]))
            continue
        if nk(k) in keys:
            warns.append(u"{}: chave repetida no TXT - descartada".format(nk(k)))
            continue
        keys.add(nk(k))
        out.append([k, t, p, origin])
    for e in out:
        if nk(e[2]) and nk(e[2]) not in keys:
            warns.append(u"{}: pai '{}' não existe - vai para o nível de cima".format(
                nk(e[0]), nk(e[2])))
            e[2] = u""
    parent = dict((nk(e[0]), nk(e[2])) for e in out)
    for e in out:                          # ciclos: sobe pelos pais até a raiz
        seen, cur = set([nk(e[0])]), nk(e[2])
        while cur:
            if cur in seen:
                warns.append(u"{}: ciclo de pais ({}) - vai para o nível de cima".format(
                    nk(e[0]), cur))
                e[2] = u""
                parent[nk(e[0])] = u""
                break
            seen.add(cur)
            cur = parent.get(cur, u"")
    return [tuple(e) for e in out], warns


def load_keynote_txt(path, expected=None):
    """Carrega o TXT como tabela de keynotes do projeto e confere o resultado
    contra as entradas geradas (chave, texto e pai). -> (ok, mensagem). ok=False
    se o Revit não confirmou a transação ou encontrou erro de sintaxe."""
    holder = {}

    def body():
        mp = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        ref = DB.ExternalResourceReference.CreateLocalResource(
            doc, DB.ExternalResourceTypes.BuiltInExternalResourceTypes.KeynoteTable,
            mp, DB.PathType.Absolute)
        res = DB.KeyBasedTreeEntriesLoadResults()
        DB.KeynoteTable.GetKeynoteTable(doc).LoadFrom(ref, res)
        holder["res"] = res
        try:
            n_syntax = len(list(res.GetFileSyntaxErrors()))
        except Exception:
            n_syntax = 0
        if n_syntax:                      # desfaz: o projeto fica com a tabela anterior
            raise Exception(u"o Revit encontrou {} erro(s) de sintaxe no arquivo - "
                            u"recarga desfeita".format(n_syntax))
    LOG.begin(u"Recarga do TXT", path)
    ok, why = run_transaction(u"Carregar TXT de Keynotes", body)
    if not ok:
        LOG.end(u"Recarga do TXT", u"FALHOU: " + why)
        return False, why
    syntax, entry_errs = [], []
    res = holder.get("res")
    try:
        syntax = [to_unicode(x) for x in res.GetFileSyntaxErrors()]
    except Exception:
        pass
    try:
        entry_errs = [to_unicode(x) for x in res.GetKeyBasedTreeEntryErrors()]
    except Exception:
        pass
    for x in syntax:
        LOG.warn(u"TXT: erro de sintaxe informado pelo Revit: {}".format(x))
    for x in entry_errs:
        LOG.warn(u"TXT: erro de entrada informado pelo Revit: {}".format(x))
    # conferência: a tabela do projeto tem as mesmas chaves/textos/pais do arquivo
    loaded, diffs = {}, []
    try:
        for e in DB.KeynoteTable.GetKeynoteTable(doc).GetKeyBasedTreeEntries():
            try:
                loaded[clean(e.Key)] = (clean(e.KeynoteText), clean(e.ParentKey))
            except Exception:
                continue
        for ent in TXT_BUILD.get("entries") or []:
            k, t, p = clean(ent[0]), clean(ent[1]), clean(ent[2])
            got = loaded.get(k)
            if got is None:
                diffs.append(u"{}: não está na tabela do projeto".format(k))
            elif u" ".join(got[0].split()) != u" ".join(t.split()) or got[1] != p:
                diffs.append(u"{}: texto/pai diferente no projeto".format(k))
    except Exception:
        LOG.exc(u"TXT: conferir tabela de keynotes do projeto")
        loaded = None
    for d in diffs[:20]:
        LOG.warn(u"TXT x projeto: " + d)
    LOG.end(u"Recarga do TXT", u"entradas no projeto: {} (esperado {}) · diferenças: {} · "
            u"sintaxe: {} · entradas com erro: {}".format(
                len(loaded) if loaded is not None else u"?", expected, len(diffs),
                len(syntax), len(entry_errs)))
    if syntax:
        return False, u"o Revit encontrou {} erro(s) de sintaxe no arquivo".format(len(syntax))
    msg = u"recarregado no projeto"
    if loaded is not None and expected is not None:
        msg += u" e conferido ({} de {} entradas{})".format(
            len(loaded), expected,
            u", {} diferença(s) - ver log".format(len(diffs)) if diffs else u", iguais ao arquivo")
    if entry_errs:
        msg += u" · {} aviso(s) do Revit".format(len(entry_errs))
    return True, msg


# ------------------------------------------------------------------
# 5h. Executa: tabela de verificação -> grava no modelo o que foi marcado
# ------------------------------------------------------------------
# (holder id, campo) -> motivo de NÃO ter sido gravado (usado no resultado)
not_written = {}
DIRS = MODE_DIRS[MODE]

VERIFY_LOG = []          # linhas da tabela 1 já gravadas (todas as rodadas)
INC_LOG = []             # linhas da tabela 2 já gravadas (todas as rodadas)


IN_STATS = {"with": 0, "without": 0, "clash": []}


def analyze_model():
    """(Re)analisa Excel x modelo e monta as 3 tabelas. Roda no início e
    depois de cada 'Atualizar descrições'. O modelo é percorrido UMA vez
    (scan_types) e o resultado é reaproveitado pelas tabelas 2/3, pela seleção
    na vista e pelo dashboard."""
    global m_rows, m_plan, m_blocked, VROWS, INC, NOTX, IN_STATS
    LOG.begin(u"Análise", u"comparação Excel x modelo")
    m_rows, m_plan, m_blocked = model_rows()
    VROWS = m_rows
    for r in VROWS:
        r["view"] = row_view(r)
        r["sel"] = r["enabled"]           # Excel -> Modelo começa marcado
    VROWS.sort(key=lambda r: (r["key"], r["category"], r["view"]["family"],
                              r["view"]["type"]))
    for i, r in enumerate(VROWS):
        r["rid"] = i
    scan = scan_types()
    INC, NOTX = collect_incomplete(set(r["key"] for r in VROWS), scan)
    for i, r in enumerate(INC):
        r["rid"] = i
        verify_incomplete(r, autofill=True)   # mostra a proposta; a linha segue desmarcada
    for i, r in enumerate(NOTX):
        r["rid"] = i
    TYPE_INST.clear()
    recs, n_bad = [], 0
    for info in scan:                     # instâncias de cada tipo (mostrar na vista)
        TYPE_INST[eid_int(info["t"].Id)] = list(info["ids"])
        try:
            recs.append(classify_type(info["t"], info))
        except Exception:
            n_bad += 1                     # tipo ilegível: fica fora (não conta como correto)
            LOG.exc(u"Dashboard: tipo {} ilegível".format(eid_int(info["t"].Id)))
    recs.sort(key=lambda x: (x["category"], x["family"], x["type"]))
    DASH["records"] = recs
    DASH["analyzedAt"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    IN_STATS = in_param_stats(scan)
    if IN_STATS["clash"]:
        LOG.warn(u"'{}' com GUID diferente do oficial (ignorado) em: {}".format(
            PARAM_IN_NAME, u", ".join(IN_STATS["clash"])))
    LOG.end(u"Análise", u"tipos analisados: {} ({} instâncias) · tabela 1: {} · tabela 2: {} · "
            u"tabela 3: {} · bloqueados: {} · '{}' em {} tipo(s), ausente em {} · "
            u"tipos ilegíveis: {}".format(
                len(scan), sum(len(i["ids"]) for i in scan), len(VROWS), len(INC), len(NOTX),
                len(m_blocked), PARAM_IN_NAME, IN_STATS["with"], IN_STATS["without"], n_bad))


def make_info_text():
    keys_with_info = [k for k in MODEL if XL["by_key"].get(k) and
                      (XL["by_key"][k]["pt"] or XL["by_key"][k]["en"])]
    txt = (u"{} keynote(s) no modelo · {} com descrição no Excel · {} fora do Excel  ·  "
           u"Tabela 1: {} linha(s) ({} tipo(s))  ·  Tabela 2: {} tipo(s) sem Keynote ou "
           u"sem Descrição  ·  Tabela 3: {} tipo(s) com Keynote fora do Excel".format(
               len(MODEL), len(keys_with_info),
               len([k for k in MODEL if not XL["by_key"].get(k)]),
               len(VROWS), sum(len(r["hs"]) for r in VROWS), len(INC), len(NOTX)))
    n_in, n_tot = IN_STATS["with"], IN_STATS["with"] + IN_STATS["without"]
    if n_tot and not n_in:
        txt += u"  ·  {}: não existe neste modelo (ignorada)".format(PARAM_IN_NAME)
    elif n_in < n_tot:
        txt += u"  ·  {}: existe em {} de {} tipo(s) (nos demais é ignorada)".format(
            PARAM_IN_NAME, n_in, n_tot)
    if IN_STATS["clash"]:
        txt += (u"  ·  Atenção: '{}' com GUID diferente do oficial em {} - não é usado"
                .format(PARAM_IN_NAME, u", ".join(IN_STATS["clash"])))
    return txt


info_text = u""


# ------------------------------------------------------------------
# 5i. "Atualizar descrições" (antes "Gravar marcados"; depois da confirmação): grava, reanalisa e ATUALIZA
#     as tabelas (a janela continua aberta). Roda via ExternalEvent.
#     "Finalizar" só fecha a interface (nada é gravado nem gerado).
# ------------------------------------------------------------------
VW = None


def inc_view(r):
    return {"category": r["category"], "family": r["family"], "type": r["name"],
            "typeId": r["id"], "count": r["count"], "hasIn": r["hasIn"],
            "key0": r["key0"], "pt0": r["pt0"], "en0": r["en0"],
            "key": clean(r["key"]), "pt": clean(r["pt"]), "en": clean(r["en"]),
            "written": list(r.get("written", [])), "check": r["check"],
            "status": r["status"]}


class _FailureCollector(DB.IFailuresPreprocessor):
    """Captura as mensagens de falha do Revit na transação: avisos são
    registrados e descartados; erros fazem a transação ser desfeita SEM abrir
    diálogo modal do Revit no meio do ExternalEvent (o motivo vai para o log e
    para o resumo da gravação)."""

    def __init__(self):
        self.errors, self.warnings = [], []

    def PreprocessFailures(self, accessor):
        has_error = False
        try:
            for m in accessor.GetFailureMessages():
                try:
                    txt = to_unicode(m.GetDescriptionText())
                except Exception:
                    txt = u"(falha sem descrição)"
                if m.GetSeverity() == DB.FailureSeverity.Warning:
                    self.warnings.append(txt)
                    try:
                        accessor.DeleteWarning(m)
                    except Exception:
                        pass
                else:
                    self.errors.append(txt)
                    has_error = True
        except Exception:
            LOG.exc(u"PreprocessFailures")
        if has_error:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


def run_transaction(name, body):
    """Executa body() numa transação e só devolve sucesso se o Revit confirmar
    o Commit. -> (ok, mensagem). Nunca abre transação aninhada."""
    try:
        if not doc.IsValidObject:
            return False, u"o documento não está mais aberto"
        if doc.IsReadOnly:
            return False, u"o documento está somente leitura"
        if doc.IsModifiable:
            return False, u"já existe uma transação aberta no documento (tente de novo)"
    except Exception:
        LOG.exc(u"run_transaction: checagem do documento")
    t = DB.Transaction(doc, name)
    fc = _FailureCollector()
    LOG.info(u"Transação '{}': Start".format(name))
    try:
        st = t.Start()
    except Exception as ex:
        LOG.exc(u"Transação '{}': Start falhou".format(name))
        return False, u"não foi possível iniciar a transação: {}".format(to_unicode(ex))
    if st != DB.TransactionStatus.Started:
        LOG.error(u"Transação '{}': Start devolveu {}".format(name, st))
        return False, u"o Revit não iniciou a transação ({})".format(to_unicode(st))
    try:                                   # depois do Start, como no exemplo da API
        fho = t.GetFailureHandlingOptions()
        fho.SetFailuresPreprocessor(fc)
        fho.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(fho)
    except Exception:
        LOG.exc(u"run_transaction: opções de falha")
    try:
        body()
    except Exception as ex:
        LOG.exc(u"Transação '{}': erro durante a gravação".format(name))
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            LOG.warn(u"Transação '{}': RollBack".format(name))
        except Exception:
            LOG.exc(u"Transação '{}': RollBack falhou".format(name))
        return False, u"erro durante a gravação (desfeita): {}".format(to_unicode(ex))
    try:
        st = t.Commit()
    except Exception as ex:
        LOG.exc(u"Transação '{}': Commit falhou".format(name))
        return False, u"o Revit recusou a gravação: {}".format(to_unicode(ex))
    for w in fc.warnings:
        LOG.warn(u"Transação '{}': aviso do Revit: {}".format(name, w))
    if st != DB.TransactionStatus.Committed:
        LOG.error(u"Transação '{}': Commit devolveu {} · erros: {}".format(
            name, st, u" | ".join(fc.errors)))
        return False, u"o Revit desfez a gravação ({}){}".format(
            to_unicode(st), u": " + u"; ".join(fc.errors[:3]) if fc.errors else u"")
    LOG.info(u"Transação '{}': Commit OK".format(name))
    return True, u""


def _set_param(p, new):
    """Grava e diz se o Revit aceitou (Parameter.Set devolve False sem
    exceção quando recusa o valor)."""
    try:
        ok = p.Set(new)
    except Exception as ex:
        return False, to_unicode(ex)
    if ok is False:
        return False, u"o Revit recusou o valor"
    return True, u""


def pending_fields(rows1, rows2):
    """Quantos campos cada gravação vai tentar: {"pt": n, "en": n, "key": n}."""
    cnt = {"pt": 0, "en": 0, "key": 0}
    for r in rows1:
        for it in r["items"]:
            if not it["reason"]:
                cnt[it["field"]] += 1
    for r in rows2:
        for f in ("key", "pt", "en"):
            if f == "en" and not r["hasIn"]:
                continue
            if clean(r[f]) and clean(r[f]) != r[f + "0"]:
                cnt[f] += 1
    return cnt


def save_run(chosen_rows, chosen_inc):
    """'Atualizar descrições' (já confirmado pelo usuário). Grava, CONFERE no
    modelo o que ficou gravado, reanalisa, atualiza as tabelas e mostra o
    resumo. Roda via ExternalEvent."""
    global MODEL
    LOG.begin(u"Gravar", u"tabela 1: {} linha(s) · tabela 2: {} linha(s)".format(
        len(chosen_rows), len(chosen_inc)))
    n_before = len(changes)
    written = []          # (Parameter, valor, registro de 'changes') para conferir
    problems = []         # (onde, motivo, "bloqueio" | "erro")
    # ignorados: linhas da tabela 1 que podiam ser gravadas e não foram marcadas
    n_ignored = len([r for r in VROWS if r["enabled"] and not r["sel"]]) + \
        len([r for r in chosen_inc if not inc_changed(r)])
    try:
        # ---- Tabela 1 (Excel -> Modelo): grava os tipos das linhas marcadas ----
        chosen_ids = set(r["rid"] for r in chosen_rows)
        for r in VROWS:
            r["status"] = (u"Bloqueado" if not r["enabled"] else
                           (u"Marcado" if r["rid"] in chosen_ids else u"Não marcado"))
        chosen_items = [it for r in chosen_rows for it in r["items"] if not it["reason"]]
        chosen_item_ids = set(id(it) for it in chosen_items)
        for it in m_plan:
            if id(it) not in chosen_item_ids:
                not_written[(it["h"]["id"], it["field"])] = (
                    u"não selecionado pelo usuário" +
                    (u" (modelo diferente do Excel)" if it["old"] else u""))
        pend1 = []

        def write1():
            for it in chosen_items:
                h, field, old, new = it["h"], it["field"], it["old"], it["new"]
                ok, why = _set_param(it["p"], new)
                if ok:
                    pend1.append((it["p"], new, {
                        "target": "model", "action": "overwrite" if old else "fill",
                        "key": h["key"], "field": it["label"], "fieldKey": field,
                        "category": h["category"], "old": old, "new": new,
                        "where": where_label(h), "id": h["id"]}))
                else:
                    problems.append((where_label(h) + u" · " + it["label"], why, u"erro"))
        if chosen_items:
            ok, msg = run_transaction(u"Atualizar descrições por Keynote", write1)
            if ok:
                written.extend(pend1)
            else:
                problems.append((u"Tabela 1", msg, u"erro"))

        # ---- Tabela 2 (tipos sem Keynote / sem Descrição) ----
        fields2 = (("key", u"Keynote", lambda t: t.get_Parameter(BIP.KEYNOTE_PARAM)),
                   ("pt", u"Description", desc_param),
                   ("en", PARAM_IN_NAME, in_param))
        for r in INC:
            r["status"], r["written"] = u"Não marcado", []
        to_write = [r for r in chosen_inc if inc_changed(r)]
        for r in chosen_inc:
            if not inc_changed(r):
                r["status"] = u"Sem alteração"
        pend2 = []

        def write2():
            for r in to_write:
                tp = r["t"]
                where = u"{}: {}{}".format(r["category"], (r["family"] + u" - ")
                                           if r["family"] else u"", r["name"])
                if owned_by_other(tp):
                    problems.append((where, u"tipo emprestado por outro usuário (workshared)", u"bloqueio"))
                    continue
                for f, label, getp in fields2:
                    new, old = clean(r[f]), r[f + "0"]
                    if not new or new == old:
                        continue
                    if f == "en" and not r["hasIn"]:
                        continue                  # Descrição IN não aplicável neste tipo
                    try:
                        p = getp(tp)
                    except Exception:
                        p = None
                    if p is None:
                        if f != "en":             # IN ausente / GUID diferente: ignora
                            problems.append((where, u"parâmetro '{}' não existe".format(label), u"bloqueio"))
                        continue
                    if param_text(p) == new:
                        r["written"].append(f)    # já gravado pela tabela 1
                        continue
                    if p.IsReadOnly:
                        problems.append((where, u"'{}' somente leitura".format(label), u"bloqueio"))
                        continue
                    ok, why = _set_param(p, new)
                    if ok:
                        pend2.append((p, new, {
                            "target": "model", "action": "overwrite" if old else "fill",
                            "key": clean(r["key"]) or NO_KEY, "field": label,
                            "fieldKey": "inc_" + f, "category": r["category"],
                            "old": old, "new": new, "where": where, "id": r["id"]}, r, f))
                    else:
                        problems.append((where + u" · " + label, why, u"erro"))
        if to_write:
            ok, msg = run_transaction(u"Completar Keynote / Descrição dos tipos", write2)
            if ok:
                for p, new, rec, r, f in pend2:
                    r["written"].append(f)
                    written.append((p, new, rec))
            else:
                problems.append((u"Tabela 2", msg, u"erro"))

        # ---- conferência: relê cada parâmetro gravado ----
        def same(a, b):                    # ignora diferença só de espaço/quebra de linha
            return u" ".join(to_unicode(a).split()) == u" ".join(to_unicode(b).split())
        by_field = defaultdict(int)
        for p, new, rec in written:
            try:
                now = param_text(p)
            except Exception:
                now = None
            if now is not None and same(now, new):
                changes.append(rec)
                by_field[rec["fieldKey"].replace(u"inc_", u"")] += 1
                LOG.info(u"Gravado e conferido: {} · {}: '{}' -> '{}'".format(
                    rec["where"], rec["field"], rec["old"], now))
            else:
                problems.append((rec["where"] + u" · " + rec["field"],
                                 u"depois de gravar o modelo mostra '{}'".format(now), u"erro"))
        for r in chosen_inc:
            if r["status"] == u"Não marcado":
                r["status"] = u"Gravado" if r.get("written") else u"Não gravado"
        for r in VROWS:
            if r["status"] == u"Marcado":
                r["status"] = u"Gravado"
        for r in chosen_rows:
            VERIFY_LOG.append(dict(r["view"], status=r["status"]))
        for r in chosen_inc:
            INC_LOG.append(inc_view(r))
        n_saved = len(changes) - n_before
        blocks = [x for x in problems if x[2] == u"bloqueio"]
        errors = [x for x in problems if x[2] == u"erro"]
        for where, why, kind in problems:
            LOG.warn(u"Não gravado ({}): {} - {}".format(kind, where, why))
        LOG.end(u"Gravar", u"{} parâmetro(s) gravado(s) e conferido(s) · {} problema(s)"
                .format(n_saved, len(problems)))

        # ---- relê o modelo e refaz as tabelas ----
        LOG.begin(u"Reanálise após gravar")
        MODEL = collect_model(use_view, opts["materials"])
        analyze_model()
        LOG.end(u"Reanálise após gravar")
        msg = (u"{} parâmetro(s) gravado(s) e conferido(s) no modelo: Description {} · "
               u"{} {} · Keynote {}.".format(n_saved, by_field["pt"], PARAM_IN_NAME,
                                            by_field["en"], by_field["key"]))
        if problems:
            msg += u" {} bloqueio(s) e {} erro(s) - veja o resumo.".format(len(blocks), len(errors))
        if VW is not None:
            VW.reload(VROWS, INC, NOTX, make_info_text(),
                      msg + u" Tabelas atualizadas - grave mais ou clique em Finalizar.")
            summary = (u"Concluído: {} parâmetro(s) gravado(s) e conferido(s) "
                       u"(Description {} · {} {} · Keynote {})\n"
                       u"Ignorados (não marcados / sem alteração): {}\n"
                       u"Bloqueados (somente leitura, emprestado, parâmetro ausente): {}\n"
                       u"Erros (Revit recusou ou desfez): {}").format(
                n_saved, by_field["pt"], PARAM_IN_NAME, by_field["en"], by_field["key"],
                n_ignored, len(blocks), len(errors))
            for title, lst in ((u"Bloqueados", blocks), (u"Erros", errors)):
                if lst:
                    summary += u"\n\n{}:\n".format(title) + u"\n".join(
                        u"• {} - {}".format(w, why) for w, why, _k in lst[:10])
                    if len(lst) > 10:
                        summary += u"\n• ... e mais {} (ver log)".format(len(lst) - 10)
            if not problems and n_saved == 0:
                summary += (u"\n\nNada foi alterado: as linhas marcadas não tinham valor "
                            u"novo para gravar (Excel vazio ou igual ao modelo).")
            _tell(VW, summary, warn=bool(problems))
    except Exception:
        LOG.exc(u"save_run")
        raise
    finally:
        try:
            if VW is not None:
                VW.b_apply.IsEnabled = True
        except Exception:
            pass


# ------------------------------------------------------------------
# 5j. Dashboard HTML de qualidade dos dados (botão "Dashboard HTML")
#   Usa o retrato gravado por analyze_model() (DASH): NÃO relê nem grava o
#   modelo. Registro = TIPO (família + tipo) usado por elementos do modelo
#   em categoria com parâmetro Keynote (mesmo universo das tabelas).
#   Critérios por registro:
#     falta ........ Keynote vazio · Description vazia · Descrição IN vazia
#     divergência .. Keynote não está no Excel · Description diferente do
#                    Excel · Descrição IN diferente do Excel (só compara quando
#                    os DOIS lados têm valor)
#     não verificado parâmetro Descrição IN inexistente no tipo, ou Excel sem
#                    descrição para comparar com a do modelo
#   Situação (uma por registro, nesta ordem): Informação faltante >
#   Divergência > Não verificado > Correto.
#   Saúde = corretos ÷ (total analisado − não verificados).
# ------------------------------------------------------------------
DASH = {"records": [], "analyzedAt": u""}


def classify_type(t, info):
    """Situação de UM tipo para o dashboard (só leitura). 'Descrição IN'
    inexistente no tipo é 'não aplicável': não conta como falta, divergência
    nem 'não verificado'."""
    key, desc, en, has_in, _missing = type_status(t)
    rec = XL["by_key"].get(key) if key else None
    missing, diverg, unver = [], [], []
    if not key:
        missing.append(u"Keynote vazio")
    if not desc:
        missing.append(u"Description vazia")
    if has_in and not en:
        missing.append(u"{} vazia".format(PARAM_IN_NAME))
    if key and rec is None:
        diverg.append(u"Keynote não está no Excel")
    if rec is not None:
        if desc and rec["pt"] and desc != rec["pt"]:
            diverg.append(u"Description diferente do Excel")
        elif desc and not rec["pt"]:
            unver.append(u"Excel sem Description para comparar")
        if has_in and en and rec["en"] and en != rec["en"]:
            diverg.append(u"{} diferente do Excel".format(PARAM_IN_NAME))
        elif has_in and en and not rec["en"]:
            unver.append(u"Excel sem {} para comparar".format(PARAM_IN_NAME))
    if missing:
        status = "missing"
    elif diverg:
        status = "diverg"
    elif unver:
        status = "unverified"
    else:
        status = "ok"
    return {"id": eid_int(t.Id), "category": info["category"], "family": family_name(t),
            "type": elem_name(t), "count": len(info["ids"]), "key": key,
            "pt": desc, "en": en, "hasIn": has_in, "inNA": not has_in,
            "inExcel": rec is not None,
            "xlRow": rec["row"] if rec is not None else None,
            "xlPt": rec["pt"] if rec is not None else u"",
            "xlEn": rec["en"] if rec is not None else u"",
            "status": status, "missing": missing, "diverg": diverg, "unverified": unver}


def build_dashboard_data():
    recs = DASH["records"]
    n = len(recs)
    by = defaultdict(int)
    for r in recs:
        by[r["status"]] += 1
    flag_missing = len([r for r in recs if r["missing"]])
    flag_diverg = len([r for r in recs if r["diverg"]])
    evaluated = n - by["unverified"]
    occ = OrderedDict()                   # ocorrência -> [grupo, quantidade]
    for r in recs:
        for grp in ("missing", "diverg", "unverified"):
            for o in r[grp]:
                occ.setdefault(o, [grp, 0])[1] += 1
    return {
        "project": to_unicode(doc.Title),
        "excel": os.path.basename(XLSX_PATH), "excelPath": XLSX_PATH,
        "sheet": XL["sheet"], "inParam": PARAM_IN_NAME,
        "analyzedAt": DASH["analyzedAt"],
        "generatedAt": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "savedThisRun": len(changes),
        "totals": {"records": n, "ok": by["ok"], "missing": by["missing"],
                   "diverg": by["diverg"], "unverified": by["unverified"],
                   "evaluated": evaluated, "flagMissing": flag_missing,
                   "flagDiverg": flag_diverg,
                   "inNA": len([r for r in recs if r.get("inNA")]),
                   "instances": sum(r["count"] for r in recs)},
        "occurrences": [{"label": k, "group": v[0], "count": v[1]} for k, v in occ.items()],
        "records": recs,
    }


def write_dashboard():
    """Grava o HTML (modelo keynotes.html + dados) e devolve o caminho."""
    template = os.path.join(HERE, "keynotes.html")
    if not os.path.isfile(template):
        raise IOError(u"Modelo do dashboard não encontrado: {}".format(template))
    with codecs.open(template, "r", encoding="utf-8") as f:
        html = brand_report(f.read())     # cores --oca-* do tema do Revit + logo
    data = json.dumps(build_dashboard_data(), ensure_ascii=False,
                      default=lambda o: to_unicode(o))
    html = html.replace("__DATA__", data.replace(u"</", u"<\\/"))
    out_path = script.get_document_data_file("keynotes_dashboard", "html")
    with codecs.open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


# ------------------------------------------------------------------
# 9. Remove as cores que as versões anteriores deixaram no modelo e abre a
#    tabela de verificação (NÃO modal)
# ------------------------------------------------------------------
# vistas em que as versões antigas pintavam (vistas de modelo); folhas, tabelas,
# legendas e vistas de desenho nunca recebiam cor
_CLEANUP_VIEW_TYPES = ("FloorPlan", "CeilingPlan", "EngineeringPlan", "AreaPlan", "Section",
                       "Elevation", "ThreeD", "Detail")
CLEANUP_BUDGET_S = 90        # tempo máximo por execução; o restante fica para a próxima


def cleanup_old_colors():
    """Remove as cores que versões ANTIGAS deste comando deixaram no modelo.
    - Só confere elementos que essas versões pintavam (instâncias de tipos com
      parâmetro Keynote, em categorias de modelo) e só em vistas de modelo.
    - Varredura completa uma vez por documento, em partes de até
      CLEANUP_BUDGET_S segundos (o progresso fica salvo); depois, só a vista
      ativa e as vistas anotadas pela versão antiga.
    - Pergunta antes de limpar. Falhas vão para o log e não param o comando."""
    LOG.begin(u"Limpeza de cores antigas")
    state_path = _colors_log_path("keynotes_cores_limpeza")
    old_log = _colors_log_path()
    state = {"complete": False, "scanned": []}
    try:
        if state_path and os.path.isfile(state_path):
            with codecs.open(state_path, "r", encoding="utf-8") as f:
                st = json.load(f)
            # arquivo da versão anterior ({"done": data}) = varredura já feita
            state["complete"] = bool(st.get("complete", "done" in st))
            state["scanned"] = list(st.get("scanned", []))
    except Exception:
        LOG.exc(u"Limpeza: leitura do estado")
    logged = set()
    if old_log and os.path.isfile(old_log):
        try:
            with codecs.open(old_log, "r", encoding="utf-8") as f:
                logged = set(int(v) for v in json.load(f).get("views", {}).keys())
        except Exception:
            LOG.exc(u"Limpeza: leitura das vistas anotadas")

    allowed = set(getattr(DB.ViewType, n) for n in _CLEANUP_VIEW_TYPES
                  if hasattr(DB.ViewType, n))
    views = [v for v in graphic_views() if v.ViewType in allowed]
    if state["complete"]:
        views = [v for v in views
                 if eid_int(v.Id) in logged or eid_int(v.Id) == eid_int(active_view.Id)]
    else:
        done = set(state["scanned"])
        views = [v for v in views if eid_int(v.Id) not in done]
    candidates = set()
    for info in scan_types():
        candidates.update(eid_int(i) for i in info["ids"])
    LOG.info(u"Limpeza: {} vista(s) para conferir · {} elemento(s) candidato(s) · "
             u"varredura completa já feita: {}".format(len(views), len(candidates),
                                                      state["complete"]))

    found, scanned, t0 = [], [], time.time()
    pb = None
    try:
        if len(views) > 3:
            pb = forms.ProgressBar(title=u"Procurando cores antigas do comando nas vistas "
                                         u"({value} de {max_value})")
            pb.__enter__()
        for n, v in enumerate(views):
            if time.time() - t0 > CLEANUP_BUDGET_S:
                LOG.warn(u"Limpeza: tempo limite atingido - {} vista(s) ficam para a "
                         u"próxima execução".format(len(views) - n))
                break
            if pb is not None:
                try:
                    pb.update_progress(n + 1, len(views))
                except Exception:
                    pass
            ids = []
            try:
                in_view = [i for i in DB.FilteredElementCollector(doc, v.Id)
                           .WhereElementIsNotElementType().ToElementIds()
                           if eid_int(i) in candidates]
                for eid in in_view:
                    try:
                        if is_old_highlight(v.GetElementOverrides(eid)):
                            ids.append(eid)
                    except Exception:
                        continue
            except Exception:
                LOG.exc(u"Limpeza: vista {} ({})".format(eid_int(v.Id), to_unicode(v.Name)))
            scanned.append(eid_int(v.Id))
            if ids:
                found.append((v, ids))
                LOG.info(u"Limpeza: vista '{}' - {} elemento(s) com cor antiga".format(
                    to_unicode(v.Name), len(ids)))
    finally:
        if pb is not None:
            try:
                pb.__exit__(None, None, None)
            except Exception:
                pass

    if not state["complete"]:
        state["scanned"] = sorted(set(state["scanned"]) | set(scanned))
        state["complete"] = len(scanned) == len(views)
    try:
        if state_path:
            with codecs.open(state_path, "w", encoding="utf-8") as f:
                json.dump({"complete": state["complete"], "scanned": state["scanned"],
                           "date": datetime.now().isoformat()}, f)
    except Exception:
        LOG.exc(u"Limpeza: gravar estado")

    if found:
        n = sum(len(ids) for _v, ids in found)
        lines = [u"  • {}: {} elemento(s)".format(to_unicode(v.Name), len(ids))
                 for v, ids in found[:12]]
        if len(found) > 12:
            lines.append(u"  • ... e mais {} vista(s)".format(len(found) - 12))
        if forms.alert(u"Remover as cores deixadas no modelo por versões anteriores deste comando?",
                       sub_msg=u"{} elemento(s) em {} vista(s):\n\n{}\n\nA sobreposição gráfica desses "
                       u"elementos volta ao padrão da vista."
                       .format(n, len(found), u"\n".join(lines)), title=ALERT_TITLE, yes=True, no=True):
            ok, bad = clear_old_highlight(found)
            LOG.info(u"Limpeza: {} elemento(s) limpos, {} falha(s)".format(ok, bad))
            if not bad and old_log and os.path.isfile(old_log):
                try:
                    os.remove(old_log)
                except Exception:
                    pass
        else:
            LOG.info(u"Limpeza: usuário não quis limpar {} elemento(s)".format(n))
    LOG.end(u"Limpeza de cores antigas", u"{} vista(s) conferida(s) em {:.1f} s · {} com cor "
            u"antiga".format(len(scanned), time.time() - t0, len(found)))


def make_keynote_txt(reload_txt=True):
    """Ao executar a 1ª aba: gera o TXT de Keynote a partir do Excel (mesma
    estrutura de sempre), grava em ANSI / Windows-1252 e recarrega no projeto.
    Arquivo existente: substituído sem perguntar, depois de copiado e conferido
    em <pasta>\\Obsoletos. Caractere fora do Windows-1252: o TXT NÃO é gravado
    e os caracteres são listados (nada vira "?"). -> mensagem para a janela.
    Só diz "recarregado" depois que o Revit confirma."""
    path = opts["txt_path"]
    LOG.begin(u"Geração do TXT", path)
    try:
        text, warns = build_keynote_txt()
    except Exception as ex:
        LOG.exc(u"Geração do TXT")
        forms.alert(u"O TXT de Keynote não foi gerado.", sub_msg=to_unicode(ex), title=ALERT_TITLE)
        return u"TXT de Keynote NÃO gerado: {}".format(to_unicode(ex))
    entries = TXT_BUILD["entries"]
    n_entries = len(entries)
    for w in warns:
        LOG.warn(u"TXT: " + w)
    LOG.end(u"Geração do TXT", u"{} entrada(s) · {} aviso(s)".format(n_entries, len(warns)))
    if n_entries == 0:
        return u"TXT de Keynote NÃO gerado: nenhuma entrada no Excel (o arquivo atual foi mantido)."
    # Windows-1252: confere ANTES de tocar em qualquer arquivo
    bad = non_cp1252_report(entries)
    if bad:
        for b in bad:
            LOG.error(u"TXT: caractere fora do Windows-1252 {} '{}' na keynote {} ({})".format(
                b["code"], b["char"], b["key"], b["origin"]))
        lines = [u"• '{}' ({}) em {} - {}{}".format(
            b["char"], b["code"], b["key"] or u"(sem chave)", b["origin"],
            u" · sugestão: troque por '{}'".format(b["hint"]) if b["hint"] else u"")
            for b in bad[:15]]
        if len(bad) > 15:
            lines.append(u"• ... e mais {} (ver log)".format(len(bad) - 15))
        forms.alert(u"O TXT de Keynote NÃO foi gravado: há caracteres que não existem no "
                    u"Windows-1252 (ANSI).",
                    sub_msg=u"Corrija no Excel, salve e use Voltar > Executar. O TXT anterior "
                            u"continua como estava.\n\n" + u"\n".join(lines),
                    title=ALERT_TITLE)
        return (u"TXT de Keynote NÃO gravado: {} caractere(s) fora do Windows-1252 no Excel "
                u"(ex.: '{}' em {}). O TXT anterior continua carregado.".format(
                    len(bad), bad[0]["char"], bad[0]["key"]))
    LOG.begin(u"Gravação do TXT")
    try:
        backup, enc = write_keynote_txt(path, text, n_entries)
    except Exception as ex:
        LOG.exc(u"Gravação do TXT")
        forms.alert(u"O TXT de Keynote não foi gravado (o arquivo anterior foi mantido).",
                    sub_msg=to_unicode(ex), title=ALERT_TITLE)
        LOG.end(u"Gravação do TXT", u"FALHOU")
        return u"TXT de Keynote NÃO gravado: {}".format(to_unicode(ex))
    LOG.end(u"Gravação do TXT")
    if reload_txt:
        ok, info = load_keynote_txt(path, n_entries)
    else:
        ok, info = True, u"NÃO recarregado (modo de diagnóstico)"
        LOG.warn(u"TXT: recarga pulada (modo de diagnóstico)")
    msg = u"TXT de Keynote gravado em {}{} e {}.{}".format(
        enc, u" · anterior em {}\\{}".format(OBSOLETE_DIR, os.path.basename(backup))
        if backup else u"", info if ok else u"NÃO recarregado no projeto: " + info,
        u"  {} aviso(s): {}".format(len(warns), u"; ".join(warns[:3])) if warns else u"")
    if not ok:
        restored = restore_txt_after_failed_reload(path, backup)
        msg += u" " + restored
        forms.alert(u"O TXT de Keynote NÃO foi recarregado no projeto - o projeto continua "
                    u"com a tabela de keynotes anterior.",
                    sub_msg=msg, title=ALERT_TITLE)
    return msg


def restore_txt_after_failed_reload(path, backup):
    """TXT e modelo não são atômicos: se o Revit não aceitou o TXT novo, o
    arquivo volta a ser o anterior (o que o projeto ainda usa). O TXT recusado
    é guardado em Obsoletos\<nome>_<data>_REJEITADO.txt. Nada é apagado.
    -> mensagem"""
    if not backup:
        return (u"Não havia TXT anterior para restaurar: o arquivo novo ficou na pasta, "
                u"mas o projeto não o carregou.")
    try:
        obs, rej = obsolete_backup_path(path)
        rej = os.path.splitext(rej)[0] + u"_REJEITADO.txt"
        File.Copy(path, rej, False)
        File.Copy(backup, path, True)
        if _bytes_of(path) != _bytes_of(backup):
            raise IOError(u"o TXT restaurado não confere com o backup")
        LOG.warn(u"TXT: recarga falhou - anterior restaurado de {} · recusado guardado em {}"
                 .format(backup, rej))
        return (u"O TXT anterior foi restaurado (igual ao projeto); o gerado ficou em "
                u"{}\{}.".format(OBSOLETE_DIR, os.path.basename(rej)))
    except Exception as ex:
        LOG.exc(u"TXT: restaurar após recarga falhar")
        return (u"ATENÇÃO: não foi possível restaurar o TXT anterior ({}). Copie manualmente "
                u"{} sobre {}.".format(_io_message(ex), backup, path))




def _step(name, fn, *args):
    """Roda uma etapa do fluxo principal com início/fim no log.
    -> (ok, resultado)"""
    LOG.begin(name)
    try:
        res = fn(*args)
    except Exception as ex:
        LOG.exc(name)
        LOG.end(name, u"FALHOU")
        forms.alert(u"Falha na etapa: {}".format(name),
                    sub_msg=u"{}\n\nDetalhes no log:\n{}".format(
                        to_unicode(ex), LOG.path or u"(log indisponível)"),
                    title=ALERT_TITLE)
        return False, None
    LOG.end(name)
    return True, res


def start_session(new_opts):
    """Lê o Excel e o modelo com as opções da 1ª aba, analisa e abre a 2ª aba.
    Roda no início do comando e de novo pelo botão Voltar (via ExternalEvent).
    Cada etapa vai para o log com início e fim. -> True se a janela abriu."""
    global opts, XLSX_PATH, XL, MODEL, FILE_TEXT, active_view, info_text, VW
    global not_written
    opts = new_opts
    XLSX_PATH = opts["xlsx"]
    active_view = doc.ActiveView
    LOG.info(u"Sessão: Excel={} · TXT={}".format(XLSX_PATH, opts["txt_path"]))
    # 1-2. Excel (leitura + cabeçalhos validados em read_excel)
    ok, xl = _step(u"1. Leitura do Excel", read_excel, XLSX_PATH)
    if not ok:
        return False
    XL = xl
    LOG.info(u"Excel: aba '{}' · {} linha(s) · {} keynote(s) distintos · {} prefixo(s) em "
             u"DADOS".format(XL.get("sheet"), len(XL.get("rows", [])), len(XL["by_key"]),
                             len(XL.get("prefix_map", {}))))
    # 3-6. TXT: geração, backup, gravação, recarga (make_keynote_txt registra cada uma)
    ok, txt_msg = _step(u"3-6. TXT de Keynote", make_keynote_txt,
                        not DIAG["skip_txt_reload"])
    if not ok:
        txt_msg = u"TXT de Keynote: falha (ver log)."
    # 7. Coleta do modelo (depois da recarga: textos do TXT relidos do projeto)
    ok, model = _step(u"7. Coleta do modelo", collect_model, use_view, opts["materials"])
    if not ok:
        return False
    MODEL = model
    FILE_TEXT = keynote_file_texts()
    LOG.info(u"Modelo: {} keynote(s) em {} tipo(s) · {} entrada(s) na tabela de keynotes".format(
        len(MODEL), sum(len(v) for v in MODEL.values()), len(FILE_TEXT)))
    not_written = {}
    del VERIFY_LOG[:]
    del INC_LOG[:]
    # 8-9. Comparação e dados das três tabelas
    ok, _r = _step(u"8-9. Comparação e tabelas", analyze_model)
    if not ok:
        return False
    info_text = make_info_text()
    if not VROWS and not INC:
        info_text += u"  ·  Nada a atualizar: o modelo já coincide com o Excel."
    # 10. Interface
    ok, win = _step(u"10. Abertura da janela", VerifyWindow, VROWS, INC, NOTX, info_text)
    if not ok:
        return False
    VW = win
    VW.viewinfo.Text = txt_msg
    try:
        VW.Show()
        LOG.info(u"Janela de verificação exibida")
    except Exception:
        LOG.exc(u"Exibir janela")
        raise
    return True


def go_back():
    """Botão Voltar: reabre a 1ª aba com as opções atuais e, se confirmada,
    analisa de novo. Cancelar na 1ª aba encerra o comando."""
    LOG.info(u"Voltar para a 1ª aba")
    new_opts = ask_options({"xlsx": opts["xlsx"], "txt_path": opts["txt_path"]}
                           if opts else None)
    if new_opts is not None:
        start_session(new_opts)


# ------------------------------------------------------------------
# 10. Modo de diagnóstico (Shift+Clique no botão)
#   Roda as etapas separadas para descobrir onde um erro acontece. Use numa
#   CÓPIA do modelo. Tudo vai para o log.
# ------------------------------------------------------------------
DIAG_OPTIONS = OrderedDict([
    ("A", u"A · Só a limpeza das cores antigas"),
    ("B", u"B · Gerar e gravar o TXT (sem recarregar no projeto)"),
    ("C", u"C · Recarregar o TXT atual no projeto (sem abrir a janela)"),
    ("D", u"D · Coletar e analisar o modelo (sem abrir a janela)"),
    ("E", u"E · Abrir a janela com dados de teste (não toca no modelo)"),
    ("F", u"F · Execução completa com log detalhado"),
    ("G", u"G · Completa SEM limpeza de cores e SEM recarga do TXT"),
    ("H", u"H · Completa com títulos de grupo simples (teste da interface)"),
    ("I", u"I · Teste de gravação da Description (grava e DESFAZ - não altera o modelo)"),
    ("L", u"Abrir a pasta dos logs"),
])


def _demo_data():
    """Linhas fictícias para testar só a interface (sem ler o modelo)."""
    global XL
    XL = {"by_key": {u"RE01": {"key": u"RE01", "pt": u"Revestimento cerâmico",
                               "en": u"Ceramic tile", "row": 5, "obs": u""}},
          "prefix_map": {u"RE": (u"REVESTIMENTO", u"MATERIAIS")}, "rows": [],
          "sheet": u"MATERIAIS", "kn_headers": [], "dados_rows": [], "cat_order": []}
    h = {"id": -1, "family": u"Basic Wall", "name": u"Parede teste", "category": u"Walls",
         "count": 0, "en_p": None, "en": None, "kind": "type", "key": u"RE01"}
    rows = [{"key": u"RE01", "h": h, "hs": [h], "items": [], "category": u"Walls",
             "pt_old": u"", "en_old": None, "pt_new": u"Revestimento cerâmico",
             "en_new": None, "over": set(), "blocked": {}, "enabled": True, "dir": "model"}]
    for r in rows:
        r["view"] = row_view(r)
        r["sel"], r["rid"] = True, 0
    base = {"t": None, "id": -2, "category": u"Floors", "family": u"Floor", "count": 0,
            "hasIn": False, "sel": False, "check": u"", "status": u""}
    inc = [dict(base, rid=0, name=u"Piso sem keynote", key0=u"", pt0=u"", en0=u"",
                key=u"", pt=u"", en=u"")]
    notx = [dict(base, rid=0, name=u"Piso ZZ", key0=u"ZZ01", pt0=u"Teste", en0=u"",
                 key=u"ZZ01", pt=u"Teste", en=u"", check=u"Keynote ZZ01 não está no Excel")]
    for r in inc:
        verify_incomplete(r, autofill=False)
    return rows, inc, notx


def write_selftest(source):
    """Teste de gravação SEM alterar o modelo: grava um valor de teste na
    Description (e na Descrição IN, se existir) de um tipo, relê e DESFAZ a
    transação (RollBack). -> lista de linhas para o relatório/log."""
    lines = [u"Executor: {}".format(source)]
    target = None
    for info in scan_types():
        try:
            p = desc_param(info["t"])
            if p is not None and not p.IsReadOnly and not owned_by_other(info["t"]):
                target = (info, p)
                break
        except Exception:
            continue
    if target is None:
        return lines + [u"Nenhum tipo com Description editável foi encontrado no modelo."]
    info, p = target
    t = info["t"]
    lines.append(u"Tipo usado: {} · {} · {} (id {})".format(
        info["category"], family_name(t), elem_name(t), eid_int(t.Id)))
    before_pt = param_text(p)
    ip = in_param(t)
    before_en = param_text(ip) if ip is not None else None
    val = u"OCA_TESTE_{}".format(datetime.now().strftime("%H%M%S"))
    tx = DB.Transaction(doc, u"Teste de gravação - Keynotes (desfeito)")
    try:
        st = tx.Start()
        lines.append(u"Transaction.Start: {}".format(st))
        ok = p.Set(val)
        lines.append(u"Description.Set: {}".format(ok))
        if ip is not None and not ip.IsReadOnly:
            lines.append(u"{}.Set: {}".format(PARAM_IN_NAME, ip.Set(val)))
        else:
            lines.append(u"{}: não aplicável neste tipo".format(PARAM_IN_NAME))
        try:
            doc.Regenerate()
        except Exception:
            pass
        now = param_text(desc_param(t))
        lines.append(u"Releitura da Description: {}".format(
            u"OK (valor de teste gravado)" if now == val else u"FALHOU - mostra '{}'".format(now)))
    except Exception as ex:
        LOG.exc(u"Teste de gravação")
        lines.append(u"ERRO: {}".format(to_unicode(ex)))
    finally:
        try:
            if tx.HasStarted() and not tx.HasEnded():
                tx.RollBack()
                lines.append(u"RollBack: feito (o modelo não foi alterado)")
        except Exception as ex:
            lines.append(u"RollBack: ERRO {}".format(to_unicode(ex)))
    after_pt = param_text(desc_param(t))
    after_en = param_text(ip) if ip is not None else None
    lines.append(u"Valores depois do teste: {}".format(
        u"iguais aos originais" if (after_pt, after_en) == (before_pt, before_en)
        else u"ATENÇÃO - diferentes dos originais"))
    for ln in lines:
        LOG.info(u"Teste de gravação: " + ln)
    return lines


def diagnostic_write_test():
    """Opção I: testa a gravação direto no comando e depois pelo mesmo caminho
    do botão 'Atualizar descrições' (ExternalEvent / Idling)."""
    env = u"Revit {} · build {} · pyRevit {}".format(
        ENV_INFO.get("revit_update"), ENV_INFO.get("revit_build"), ENV_INFO.get("pyrevit"))
    direct = write_selftest(u"direto no comando")
    t0 = time.time()

    def via_event():
        lines = write_selftest(u"ação externa (mesmo caminho do botão)")
        lines.insert(1, u"Tempo até executar: {:.1f} s".format(time.time() - t0))
        forms.alert(u"Teste de gravação - parte 2 de 2 (caminho do botão)",
                    sub_msg=u"{}\n\n{}\n\nLog: {}".format(env, u"\n".join(lines), LOG.path),
                    title=ALERT_TITLE)
    queued = run_in_revit(via_event, u"Teste de gravação")
    forms.alert(u"Teste de gravação - parte 1 de 2 (direto no comando)",
                sub_msg=u"{}\n\n{}\n\n{}".format(
                    env, u"\n".join(direct),
                    u"A parte 2 abre sozinha em seguida. Se ela NÃO aparecer em até 10 s, "
                    u"o Revit não está executando ações externas - me envie o log."
                    if queued else u"A parte 2 NÃO pôde ser enfileirada (ver log)."),
                title=ALERT_TITLE)


def diagnostic_mode():
    LOG.info(u"Modo de diagnóstico")
    pick = forms.CommandSwitchWindow.show(
        list(DIAG_OPTIONS.values()),
        message=u"Modo de diagnóstico (versão {}) - use numa CÓPIA do modelo. Log: {}".format(
            SCRIPT_VERSION, LOG.path or u"(indisponível)"))
    if not pick:
        return
    code = [k for k, v in DIAG_OPTIONS.items() if v == pick][0]
    LOG.info(u"Diagnóstico: opção {}".format(pick))
    global opts, XLSX_PATH, XL, MODEL, FILE_TEXT, VW
    if code == "L":
        folder = os.path.dirname(LOG.path) if LOG.path else None
        if folder and not open_in_browser(folder):
            forms.alert(u"Pasta dos logs:", sub_msg=folder, title=ALERT_TITLE)
        return
    if code == "I":
        _step(u"I. Teste de gravação", diagnostic_write_test)
        return
    if code == "A":
        ok, _r = _step(u"A. Limpeza de cores antigas", cleanup_old_colors)
        forms.alert(u"Diagnóstico A {}.".format(u"concluído" if ok else u"FALHOU"),
                    sub_msg=u"Detalhes (vistas conferidas, tempo, elementos) no log:\n{}".format(
                        LOG.path), title=ALERT_TITLE)
    elif code == "E":
        rows, inc, notx = _demo_data()
        VW = VerifyWindow(rows, inc, notx, u"DADOS DE TESTE - nada é gravado no modelo.",
                          demo=True)
        VW.Show()
        return
    elif code in ("F", "G", "H"):
        DIAG["skip_cleanup"] = code == "G"
        DIAG["skip_txt_reload"] = code == "G"
        DIAG["simple_group_rows"] = code == "H"
        main()
        return
    else:
        o = ask_options()
        if o is None:
            return
        opts, XLSX_PATH = o, o["xlsx"]
        if code in ("B", "D"):
            ok, xl = _step(u"Leitura do Excel", read_excel, XLSX_PATH)
            if not ok:
                return
            XL = xl
        if code == "B":
            ok, msg = _step(u"B. TXT sem recarga", make_keynote_txt, False)
        elif code == "C":
            ok, res = _step(u"C. Recarga do TXT", load_keynote_txt, o["txt_path"])
            msg = (u"Carregado: " if res and res[0] else u"NÃO carregado: ") + (
                res[1] if res else u"")
        else:
            ok, model = _step(u"D. Coleta do modelo", collect_model, use_view, False)
            if ok:
                MODEL = model
                FILE_TEXT = keynote_file_texts()
                ok, _r = _step(u"D. Análise", analyze_model)
            msg = make_info_text() if ok else u"falhou"
        forms.alert(u"Diagnóstico {} concluído.".format(code),
                    sub_msg=u"{}\n\nLog: {}".format(msg, LOG.path), title=ALERT_TITLE)
    LOG.info(u"Diagnóstico {} concluído".format(code))


def main():
    first = ask_options()
    if first is None:
        LOG.info(u"Cancelado na 1ª aba")
        return
    if not DIAG["skip_cleanup"]:
        try:
            cleanup_old_colors()
        except Exception:
            LOG.exc(u"Limpeza de cores antigas")
    else:
        LOG.warn(u"Limpeza de cores antigas pulada (diagnóstico)")
    start_session(first)


ENV_INFO = revit_env_info()
LOG.info(u"===== Início: {} versão {} ({}) · log: {} =====".format(
    TOOL_TITLE, SCRIPT_VERSION, u"diagnóstico" if CONFIG_MODE else u"normal", LOG.path))
try:
    if CONFIG_MODE:
        diagnostic_mode()
    else:
        main()
except SystemExit:
    raise
except Exception as _ex:
    LOG.exc(u"Fluxo principal")
    forms.alert(u"O comando encontrou um erro e foi interrompido.",
                sub_msg=u"{}\n\nDetalhes no log:\n{}".format(to_unicode(_ex), LOG.path),
                title=ALERT_TITLE)
