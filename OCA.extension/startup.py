# -*- coding: utf-8 -*-
import traceback

try:
    import os
    import io
    import clr
    
    clr.AddReference('AdWindows')
    clr.AddReference('PresentationCore')
    import Autodesk.Windows as adwin
    import System.Windows.Media as media
    from Autodesk.Revit.UI import TaskDialog, UITheme, UIThemeManager
    
    DEBUG_TXT = "--- DIAGNOSTICO OCA ---\n\n"
    
    # 1. OBTER PASTA
    ext_dir = os.path.join(os.environ.get('APPDATA'), r"pyRevit\Extensions\OCA.extension")
    try:
        if os.path.exists(os.path.dirname(__file__)):
            ext_dir = os.path.dirname(__file__)
    except:
        pass
        
    DEBUG_TXT += "1. Encontrou a pasta da extensao? {}\n".format(os.path.exists(ext_dir))
    
    cores_paineis = {}

    # 2. LER YAMLS
    if os.path.exists(ext_dir):
        for root, dirs, files in os.walk(ext_dir):
            for dir_name in dirs:
                if dir_name.endswith('.panel'):
                    nome_painel = dir_name.replace('.panel', '').strip().lower()
                    yaml_path = os.path.join(root, dir_name, 'bundle.yaml')
                    
                    if os.path.exists(yaml_path):
                        cor_clara, cor_escura = None, None
                        with io.open(yaml_path, 'r', encoding='utf-8', errors='ignore') as f:
                            linhas = f.readlines()
                            
                        modo_atual = None
                        for linha in linhas:
                            l = linha.strip().replace('\x00', '') 
                            if l.startswith("background:"): modo_atual = "light"
                            elif l.startswith("background_dark:"): modo_atual = "dark"
                            elif l.startswith("title:") and modo_atual:
                                cor = l.replace("title:", "").replace('"', '').replace("'", "").strip()
                                if len(cor) == 7: cor = cor.replace("#", "#FF")
                                if modo_atual == "light": cor_clara = cor
                                elif modo_atual == "dark": cor_escura = cor
                                modo_atual = None
                                
                        if cor_clara or cor_escura:
                            cores_paineis[nome_painel] = {
                                "light": cor_clara or cor_escura, 
                                "dark": cor_escura or cor_clara
                            }
                            DEBUG_TXT += "2. Lida cor do painel [{}]: Claro={} | Escuro={}\n".format(nome_painel, cor_clara, cor_escura)

    # 3. FUNCAO DE ATUALIZACAO DA RIBBON
    def atualizar_ribbon(is_dark):
        paineis_pintados = 0
        for tab in adwin.ComponentManager.Ribbon.Tabs:
            if tab.Id and "OCA" in tab.Id:
                for panel in tab.Panels:
                    if panel.Source and panel.Source.Title:
                        titulo = panel.Source.Title.replace('\n', ' ').strip().lower()
                        if titulo in cores_paineis:
                            cores = cores_paineis[titulo]
                            cor_alvo = cores["dark"] if is_dark else cores["light"]
                            panel.CustomPanelTitleBarBackground = media.SolidColorBrush(media.ColorConverter.ConvertFromString(cor_alvo))
                            paineis_pintados += 1
        return paineis_pintados

    # 4. ESPERAR O REVIT CARREGAR
    def no_arranque_concluido(sender, args):
        global DEBUG_TXT
        __revit__.Idling -= no_arranque_concluido
        try:
            # Verifica o tema nativamente pelo Revit 2024+
            tema_escuro_ativo = (UIThemeManager.CurrentTheme == UITheme.Dark)
            
            qtd = atualizar_ribbon(tema_escuro_ativo)
            DEBUG_TXT += "\n3. Paineis pintados com sucesso na Ribbon: {}".format(qtd)
            
            # Se quiser parar de ver a mensagem de diagnostico, coloque um # na frente da linha abaixo:
            # TaskDialog.Show("Diagnostico OCA", DEBUG_TXT)
        except Exception as e:
            TaskDialog.Show("Erro Aplicar Cor OCA", str(e))

    __revit__.Idling += no_arranque_concluido

    # 5. ATUALIZAR QUANDO MUDA TEMA
    def ao_mudar_tema(sender, args):
        atualizar_ribbon(args.Theme == UITheme.Dark)

    __revit__.ThemeChanged += ao_mudar_tema

except Exception as e:
    TaskDialog.Show("Erro Fatal OCA", traceback.format_exc())