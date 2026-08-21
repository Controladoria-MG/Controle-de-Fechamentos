import json
import os
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import win32api
import win32con
import win32gui
import win32process
from dotenv import load_dotenv
from pywinauto import Application
from pywinauto.keyboard import send_keys

RAIZ = Path(__file__).parent.parent
load_dotenv(RAIZ / ".env")

MGAPPS_EXE = r"C:\Program Files (x86)\MGApps\MGApps.Presentation.WpfApp.exe"
PASTA_DESTINO = RAIZ / "data" / "radar_fiscal"
ARQUIVO_SAIDA = PASTA_DESTINO / "radar_fiscal.xlsx"
ARQUIVO_MERCADOS = PASTA_DESTINO / "planilha_mercados.xlsx"
ARQUIVO_RESUMO = PASTA_DESTINO / "resumo.xlsx"
ARQUIVO_STATUS = PASTA_DESTINO / "status.json"
ARQUIVO_DADOS_PORTAL = PASTA_DESTINO / "radar_fiscal_dados.json"

PASTA_DOWNLOADS = Path.home() / "Downloads"

# Cópia de rede que alimenta a Power Query de "Resumo Radar Fiscal.xlsx"
# (arquivo do usuário fora deste repo, em Y:) — ele referencia essas duas
# planilhas diretamente e as atualiza sozinho ao ser aberto/atualizado, então
# só precisamos manter os arquivos brutos aqui em dia. Ver [[project_radar_fiscal]].
PASTA_REDE = Path(r"Y:\GERÊNCIA\Gerencia Operacional\CTR\Juan\Controle de Fechamentos")
ARQUIVO_REDE_RADAR = PASTA_REDE / "Radar Fiscal" / "radar_fiscal.xlsx"
ARQUIVO_REDE_MERCADOS = PASTA_REDE / "Planilha de Mercados" / "planilha_mercados.xlsx"

UNIDADES_MERCADOS = {"SP", "GOIAS"}

# Mesma lógica de limpeza usada na consulta Power Query "Radar Fiscal" do
# arquivo de referência do usuário (Resumo Radar Fiscal.xlsx).
MAPA_STATUS = {
    "fechado": "Fechado",
    "bloqueado": "Bloqueado",
    "simulando": "Simulando",
    "OK": "Com o GC",
}
MAPA_REGIME = {
    "Federal - Lucro Presumido": "Lucro Presumido",
    "Federal - L Real -Trimestral": "Lucro Real",
    "Federal - SN": "Simples Nacional",
    "Federal - MEI": "MEI",
    "Federal - Lucro Real - Anual": "Lucro Real",
    "Federal - L.Real - Mensal": "Lucro Real",
    "Federal - Imune": "Imune",
}
# O Sistema de Analise exporta "Gerente de Contas" com grafia inconsistente
# (ex.: "RODRIGO SILVA REGO", "tiago", "BRUNO MOTA") — normaliza pro nome
# padronizado que o usuário definiu. Chave = nome bruto em maiúsculas/sem
# espaços nas pontas, pra pegar qualquer variação de caixa.
MAPA_GERENTES = {
    "BEATRIZ RANGEL": "Beatriz Rangel",
    "DANIEL OLIVEIRA": "Daniel Oliveira",
    "ROGERIO EGIDIO": "Rogerio Egidio",
    "ANDRE LOPES COELHO": "Andre Lopes Coelho",
    "RODRIGO SILVA REGO": "Rodrigo Silva Rego",
    "SIMONE MARQUES": "Simone Marques",
    "RICARDO FERREIRA": "Ricardo Ferreira",
    "CRISTIANE SIMOES": "Cristiane Simoes",
    "BRUNO CONTIERI": "Bruno Contieri",
    "VANDER LUCIO DA SILVA": "Vander Lucio da Silva",
    "DIEGO APARECIDO CORREIA PAZ DE LIMA": "Diego Aparecido Correia Paz de Lima",
    "KARINE GUSMÃO": "Karine Gusmão",
    "RAFAELLA MASSUIA": "Rafaella Massuia",
    "LUIZ CARLOS FERREIRA": "Luiz Carlos Ferreira",
    "TIAGO RODRIGUES": "Tiago Rodrigues",
    "TIAGO": "Tiago Rodrigues",
    "MAICON COSTA": "Maicon Costa",
    "BRUNO MOTA": "Bruno Mota",
}
# Colunas da Planilha de Mercados trazidas para o resumo (mesmas 6 da consulta
# original), renomeadas com o prefixo "Planilha de Mercados." como no Power Query.
COLUNAS_MERCADOS_TRAZIDAS = [
    "ÍA", "DESC. REMESSAS", "Data Comentário Operação",
    "EFV ATUAL Data", "AÇÃO GERENTE", "Data Comentário Gerência",
]

COLUNAS_PORTAL = [
    "IdCorporativo", "Nome", "Grupo", "Unidade", "Segmento", "Gerente de Contas",
    "Status", "DeptoFiscal", "RegimeApuracao", "EmpSemMovto",
    "Documentação", "DiasSemAtualizacao", "Planilha de Mercados.AÇÃO GERENTE",
    "Planilha de Mercados.DESC. REMESSAS", "DataConfirmacao",
]

USUARIO = os.getenv("USUARIO", "")
SENHA = os.getenv("SENHA", "")

TIMEOUT_ELEMENTO = 30
TIMEOUT_EXPORTACAO = 120

# Esta automação roda quase inteiramente "em segundo plano": os controles do
# Sistema de Analise são WinForms nativos e respondem a UI Automation
# (.invoke()) e a mensagens de janela (PostMessage), sem precisar mover o
# mouse ou digitar de verdade — o script pode rodar sem tomar o controle do
# computador e mesmo com outras janelas em foco.
#
# A única exceção é abrir o app "Sistema de Analise" a partir do launcher
# MG Apps: esse launcher é WPF (não WinForms) e seus "tiles" não respondem a
# UI Automation nem a mensagens de janela — só a um clique de mouse real
# (testado: .invoke()/DoDefaultAction não disparam a ação). Por isso, e só
# nessa etapa, o script usa click_input() de verdade.


# ── Utilitários de janela ────────────────────────────────────────────────────

def _janela_por_titulo(titulo: str, exato: bool = True, timeout: float = TIMEOUT_ELEMENTO):
    """Localiza uma janela top-level pelo título usando win32gui diretamente.
    A enumeração via Desktop(backend='uia').windows() se mostrou instável para
    diálogos padrão do Windows (classe #32770, ex.: confirmações e 'Salvar
    como'): às vezes não os retorna mesmo com a janela visível na tela."""
    fim = time.time() + timeout
    while time.time() < fim:
        encontrados = []

        def cb(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                texto = win32gui.GetWindowText(hwnd)
                if (texto == titulo) if exato else texto.startswith(titulo):
                    encontrados.append(hwnd)

        win32gui.EnumWindows(cb, None)
        if encontrados:
            return encontrados[-1]
        time.sleep(0.5)
    raise TimeoutError(f"Janela '{titulo}' não apareceu em {timeout}s.")


def _conectar(handle):
    app = Application(backend="uia").connect(handle=handle)
    return app.window(handle=handle)


def _ocultar_janela(hwnd: int):
    """'Estaciona' a janela bem longe de qualquer monitor (não minimiza).
    A automação dessas janelas (WinForms) não depende delas estarem visíveis
    ou em primeiro plano — .invoke() e PostMessage funcionam do mesmo jeito —
    então isso só tira a janela da vista do usuário sem afetar o robô."""
    win32gui.SetWindowPos(
        hwnd, 0, -32000, -32000, 0, 0,
        win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE,
    )


def _invocar(janela, **criterios):
    """Aciona um botão/item de menu via UI Automation (InvokePattern), sem
    mover o mouse nem precisar que a janela esteja em primeiro plano."""
    janela.child_window(**criterios).invoke()


def _localizar_popup_menu(timeout: float = 5) -> int:
    """Menus suspensos (ToolStripDropDownMenu) não aparecem na árvore UIA da
    janela que os originou — surgem como uma janela top-level separada
    (classe 'WindowsForms10.Window.20808...'). Retorna o handle dessa janela."""
    fim = time.time() + timeout
    while time.time() < fim:
        achado = []

        def cb(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                cls = win32gui.GetClassName(hwnd)
                if "WindowsForms10.Window.20808" in cls:
                    achado.append(hwnd)

        win32gui.EnumWindows(cb, None)
        if achado:
            return achado[-1]
        time.sleep(0.2)
    raise TimeoutError("Menu suspenso não apareceu a tempo.")


def _clicar_ponto_por_mensagem(hwnd: int, x_tela: int, y_tela: int):
    """Clica num ponto de tela específico de uma janela via mensagens
    (PostMessage), sem mover o cursor do mouse. Usa PostMessage (não
    SendMessage) porque o clique pode abrir um diálogo modal em seguida —
    SendMessage travaria esperando esse diálogo fechar."""
    x_cli, y_cli = win32gui.ScreenToClient(hwnd, (x_tela, y_tela))
    lparam = win32api.MAKELONG(x_cli, y_cli)
    win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lparam)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    time.sleep(0.1)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
    time.sleep(0.3)


def _clicar_item_do_menu(hwnd_popup: int, indice: int):
    """Clica em um item de um ToolStripDropDownMenu por posição (os itens não
    são expostos via UI Automation — são owner-drawn)."""
    l, t, r, _b = win32gui.GetWindowRect(hwnd_popup)
    largura = r - l
    x_tela = l + int(largura * 0.4)
    y_tela = t + 18 + 22 * indice  # itens empilhados, ~22px de altura, 1º item em y≈18
    _clicar_ponto_por_mensagem(hwnd_popup, x_tela, y_tela)


def _forcar_primeiro_plano(hwnd: int) -> bool:
    """Tenta trazer a janela para primeiro plano de fato. O Windows bloqueia
    SetForegroundWindow vindo de um processo que não é o 'ativo' (proteção
    contra apps roubarem o foco) — o truque padrão é anexar temporariamente
    a fila de input da thread atual à da janela em primeiro plano."""
    if win32gui.GetForegroundWindow() == hwnd:
        return True
    fg_hwnd = win32gui.GetForegroundWindow()
    fg_thread = win32process.GetWindowThreadProcessId(fg_hwnd)[0] if fg_hwnd else 0
    current_thread = win32api.GetCurrentThreadId()
    anexado = False
    try:
        if fg_thread and fg_thread != current_thread:
            win32process.AttachThreadInput(current_thread, fg_thread, True)
            anexado = True
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.BringWindowToTop(hwnd)
        win32gui.SetForegroundWindow(hwnd)
    finally:
        if anexado:
            win32process.AttachThreadInput(current_thread, fg_thread, False)
    return win32gui.GetForegroundWindow() == hwnd


def _clicar_com_seguranca(elemento, hwnd_janela: int, log=None):
    """Clique real (mouse) — só usado para os 'tiles' do launcher MG Apps
    (WPF), que não respondem a UI Automation nem a mensagens de janela.
    Antes de clicar, confirma que a janela realmente está em primeiro plano E
    que o ponto de clique realmente pertence a ela — para NUNCA clicar cegamente
    em outro app que por acaso esteja por cima (ex.: se o usuário estava
    usando o mouse/outra janela no momento)."""
    if not _forcar_primeiro_plano(hwnd_janela):
        raise RuntimeError(
            "Não foi possível trazer a janela do MGApps para primeiro plano "
            "(Windows bloqueou a troca de foco). Abortando para não clicar "
            "em outra janela por engano."
        )
    rect = elemento.rectangle()
    x, y = rect.mid_point()
    hwnd_no_ponto = win32gui.WindowFromPoint((x, y))
    _, pid_no_ponto = win32process.GetWindowThreadProcessId(hwnd_no_ponto)
    _, pid_janela = win32process.GetWindowThreadProcessId(hwnd_janela)
    if pid_no_ponto != pid_janela:
        raise RuntimeError(
            "O ponto de clique não pertence à janela do MGApps (outra janela "
            "está sobrepondo). Abortando para não clicar em outro app."
        )
    elemento.click_input()


def _fechar_dialogo_se_existir(titulo: str, botao: str, timeout: float = 3) -> bool:
    try:
        handle = _janela_por_titulo(titulo, timeout=timeout)
    except TimeoutError:
        return False
    _invocar(_conectar(handle), title=botao, control_type="Button")
    return True


# ── Passos do fluxo ──────────────────────────────────────────────────────────

def _fechar_mgapps_se_existir(log=None):
    """Fecha a janela 'MG Apps' se já estiver aberta (execução anterior que
    não chegou a fechar, crash, ou o usuário abriu manualmente) — a
    automação sempre começa do zero em vez de reaproveitar uma janela com
    estado desconhecido (ex.: estacionada fora da tela por _ocultar_janela
    numa execução anterior que travou no meio, ou numa tela/diálogo
    inesperado), que foi a origem de falhas em cascata vistas ao vivo
    (2026-08-21: reaproveitar um MGApps 'preso' de uma execução que crashou
    fez o clique no tile 'Sistema de Analise' falhar repetidamente). Mesmo
    padrão que _fechar_analise_balanco_se_existir() usa no projeto
    Analise-de-Balanco."""
    try:
        handle = _janela_por_titulo("MG Apps", timeout=2)
    except TimeoutError:
        return
    _log("MGApps já estava aberto — fechando para começar do zero...", log)
    try:
        _conectar(handle).close()
    except Exception:
        pass
    time.sleep(1)


def _abrir_mgapps(log=None):
    _fechar_mgapps_se_existir(log)
    _log("Abrindo MGApps...", log)
    subprocess.Popen([MGAPPS_EXE])
    handle = _janela_por_titulo("MG Apps", timeout=20)
    mgapps = _conectar(handle)
    mgapps.restore()
    mgapps.set_focus()
    time.sleep(1)
    return mgapps


def _garantir_tela_sistemas(mgapps, log=None):
    """A janela do MGApps guarda o estado da última tela usada (ex.: o
    painel 'Contatos', aberto pelo ícone de telefone). Se não estiver na
    tela 'Sistemas', digitar na busca vai cair no campo errado (busca de
    colaborador em vez de busca de sistema). Fecha o painel 'Contatos' se
    ele estiver aberto."""
    contatos = mgapps.child_window(title="Contatos", control_type="Text")
    if contatos.exists():
        _log("MGApps estava na tela 'Contatos', voltando para 'Sistemas'...", log)
        voltar = mgapps.child_window(auto_id="PART_Header", control_type="Custom").children(control_type="Button")[0]
        _clicar_com_seguranca(voltar, mgapps.handle, log)
        time.sleep(1)


def _abrir_sistema_analise(mgapps, log=None):
    _garantir_tela_sistemas(mgapps, log)
    _log("Procurando 'Sistema de Analise'...", log)
    campo_busca = mgapps.child_window(control_type="Edit")
    campo_busca.set_edit_text("")
    campo_busca.set_edit_text("Sistema de Analise")

    # Numa instância recém-aberta do MGApps a lista de sistemas ainda está
    # carregando ("Carregando sistemas...") por alguns segundos — espera o
    # tile realmente existir em vez de um sleep fixo, que falha nesse caso.
    tile = mgapps.child_window(title="Sistema de Analise", control_type="Text")
    tile.wait("exists", timeout=TIMEOUT_ELEMENTO)
    time.sleep(0.5)

    # Único clique real do fluxo — ver nota no topo do arquivo.
    for tentativa in range(1, 4):
        _clicar_com_seguranca(tile, mgapps.handle, log)
        try:
            _janela_por_titulo("Login", timeout=8)
            return
        except TimeoutError:
            _log(f"Clique não abriu o Login (tentativa {tentativa}/3), tentando de novo...", log)
    raise TimeoutError("Não foi possível abrir 'Sistema de Analise' após 3 tentativas.")


def _login_sistema_analise(log=None):
    _log("Autenticando no Sistema de Analise...", log)
    login_handle = _janela_por_titulo("Login", timeout=TIMEOUT_ELEMENTO)
    _ocultar_janela(login_handle)
    login = _conectar(login_handle)
    login.child_window(auto_id="UsernameTextBox", control_type="Edit").set_edit_text(USUARIO)
    login.child_window(auto_id="PasswordTextBox", control_type="Edit").set_edit_text(SENHA)
    _invocar(login, auto_id="OK", control_type="Button")


def _abrir_importacao_fiscal(log=None):
    menu_handle = _janela_por_titulo("Menu", timeout=TIMEOUT_ELEMENTO)
    _ocultar_janela(menu_handle)
    menu = _conectar(menu_handle)
    _log("Abrindo Importação Fiscal...", log)
    _invocar(menu, auto_id="btnImportacaoFiscal", control_type="Button")
    fiscal_handle = _janela_por_titulo("Departamento Fiscal", exato=False, timeout=TIMEOUT_ELEMENTO)
    _ocultar_janela(fiscal_handle)
    return _conectar(fiscal_handle)


def _abrir_radar_fiscal(fiscal, log=None):
    _log("Abrindo Relatório > Radar Fiscal...", log)
    _invocar(fiscal, title="Relatório", control_type="MenuItem")
    popup = _localizar_popup_menu()
    _clicar_item_do_menu(popup, indice=2)  # "Radar Fiscal" é o 3º item (índice 2)
    radar_handle = _janela_por_titulo("Radar", timeout=TIMEOUT_ELEMENTO)
    _ocultar_janela(radar_handle)
    return _conectar(radar_handle)


def _filtrar(radar, log=None):
    _log("Filtrando (Ano/Mês já vêm preenchidos automaticamente)...", log)
    _invocar(radar, auto_id="btnFiltrar", control_type="Button")
    time.sleep(4)


def _exportar_excel(radar, caminho_destino: Path, log=None):
    _log("Exportando para Excel...", log)
    PASTA_DESTINO.mkdir(parents=True, exist_ok=True)
    if caminho_destino.exists():
        caminho_destino.unlink()

    # O menu "Arquivo" às vezes não abre o popup a tempo no primeiro invoke
    # (visto ao vivo em 2026-08-21, reproduzido 2 de 2 vezes) — retry com ESC
    # entre tentativas pra fechar qualquer dropdown que tenha ficado parcialmente
    # aberto, mesmo padrão já usado em _exportar() no projeto Analise-de-Balanco
    # (backend/radar_fechamento.py).
    popup = None
    for tentativa in range(1, 4):
        send_keys("{ESC}")
        time.sleep(0.3)
        _invocar(radar, title="Arquivo", control_type="MenuItem")
        try:
            popup = _localizar_popup_menu()
            break
        except TimeoutError:
            _log(f"Menu 'Arquivo' não abriu (tentativa {tentativa}/3), tentando de novo...", log)
    if popup is None:
        raise TimeoutError("Não foi possível abrir o menu 'Arquivo' após 3 tentativas.")
    _clicar_item_do_menu(popup, indice=0)  # "Exportar para Excel" é o 1º item (índice 0)

    confirmar_handle = _janela_por_titulo("Sistema de Analise", timeout=TIMEOUT_EXPORTACAO)
    _invocar(_conectar(confirmar_handle), title="OK", control_type="Button")

    # O diálogo "Salvar como" é um diálogo nativo do Windows (não WinForms comum)
    # — mover ele com SetWindowPos trava a conclusão do salvamento. Fica visível
    # mesmo (é rápido, poucos segundos).
    salvar_handle = _janela_por_titulo("Salvar como", timeout=TIMEOUT_ELEMENTO)
    salvar = _conectar(salvar_handle)
    campo_nome = salvar.child_window(auto_id="FileNameControlHost", control_type="ComboBox")
    campo_nome.child_window(control_type="Edit").set_edit_text(str(caminho_destino))
    _invocar(salvar, title="Salvar", auto_id="1", control_type="Button")

    # Se já existir um arquivo com o mesmo nome, o Windows pergunta se deseja substituir
    _fechar_dialogo_se_existir("Confirmar Salvar Como", "Sim", timeout=3)

    sucesso_handle = _janela_por_titulo("Sistema de Analise", timeout=TIMEOUT_EXPORTACAO)
    _invocar(_conectar(sucesso_handle), title="OK", control_type="Button")

    fim = time.time() + TIMEOUT_EXPORTACAO
    while time.time() < fim:
        if caminho_destino.exists():
            return
        time.sleep(1)
    raise RuntimeError(f"Exportação não gerou o arquivo em {TIMEOUT_EXPORTACAO}s.")


def _copiar_para_rede(origem: Path, destino: Path, log=None):
    """Atualiza a cópia de rede (Y:) que a Power Query de 'Resumo Radar
    Fiscal.xlsx' lê. Não derruba a execução se a rede estiver indisponível —
    só avisa, já que o resto do pipeline (resumo.xlsx + portal) não depende
    disso."""
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(origem, destino)
        _log(f"Cópia de rede atualizada: {destino}", log)
    except OSError as e:
        _log(f"AVISO: não foi possível atualizar a cópia de rede {destino} ({e}).", log)


def _fechar_janela_se_existir(titulo: str, exato: bool = True, timeout: float = 2):
    try:
        handle = _janela_por_titulo(titulo, exato=exato, timeout=timeout)
        _conectar(handle).close()
    except Exception:
        pass


def _abrir_controle_status(fiscal, log=None):
    _log("Abrindo Análise Fechamento Escrita Fiscal > Controle de Status...", log)
    _invocar(fiscal, title="Análise Fechamento Escrita Fiscal", control_type="MenuItem")
    popup = _localizar_popup_menu()
    _clicar_item_do_menu(popup, indice=1)  # "Controle de Status" é o 2º item (índice 1)
    handle = _janela_por_titulo("Controle de Análise Fechamento Escrita Fiscal", timeout=TIMEOUT_ELEMENTO)
    _ocultar_janela(handle)
    return _conectar(handle)


def _limpar_filtro_mercados(janela, log=None):
    _invocar(janela, title="Limpar", control_type="Button")


def _tabela_por_cabecalho(janela, cabecalho: str):
    """As duas listas de checkbox (Unidade, Segmento) são DataGridView sem
    nome distintivo — identificadas pelo texto do cabeçalho da 2ª coluna
    (os cabeçalhos ficam dentro de uma linha 'Linha Superior', não como
    filhos diretos da tabela — por isso usa descendants, não children)."""
    for tabela in janela.descendants(control_type="Table"):
        cabecalhos = [h.window_text() for h in tabela.descendants(control_type="Header")]
        if cabecalho in cabecalhos:
            return tabela
    raise RuntimeError(f"Tabela com cabeçalho '{cabecalho}' não encontrada.")


def _marcar_unidades(janela, nomes: set, log=None):
    """Marca as checkboxes da lista 'Unidade' cujo texto está em `nomes`.
    Busca por texto (não por índice de linha) porque a grade é reaproveitada
    ao rolar. O clique é por mensagem (PostMessage) na própria janela — a
    mesma técnica usada nos menus suspensos."""
    tabela = _tabela_por_cabecalho(janela, "Unidade")
    faltando = set(nomes)
    for row in tabela.children():
        if row.element_info.control_type != "Custom" or not faltando:
            continue
        celulas = row.children()
        if len(celulas) < 2:
            continue
        checkbox, edit = celulas[0], celulas[1]
        valor = edit.legacy_properties().get("Value")
        if valor in faltando:
            if checkbox.get_toggle_state() == 0:
                rect = checkbox.rectangle()
                x, y = rect.mid_point()
                _clicar_ponto_por_mensagem(janela.handle, x, y)
            faltando.discard(valor)
    if faltando:
        raise RuntimeError(f"Unidade(s) não encontrada(s) na lista para marcar: {faltando}")
    _log(f"Unidades marcadas: {sorted(nomes)}", log)


def _filtrar_mercados(janela, log=None):
    _log("Filtrando Planilha de Mercados...", log)
    _invocar(janela, title="Filtrar", control_type="Button")
    time.sleep(3)


def _exportar_planilha_mercados(janela, caminho_destino: Path, log=None):
    _log("Exportando Planilha de Mercados...", log)
    PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
    antes = {p.name for p in PASTA_DOWNLOADS.glob("Fechamento Escrita Fiscal*.xlsx")}

    _invocar(janela, title="Relatório", control_type="MenuItem")
    popup = _localizar_popup_menu()
    _clicar_item_do_menu(popup, indice=0)  # "Exportar Excel" é o único item

    pasta_handle = _janela_por_titulo("Procurar Pasta", timeout=TIMEOUT_ELEMENTO)
    pasta_dlg = _conectar(pasta_handle)
    item = pasta_dlg.child_window(title="Downloads", control_type="TreeItem")
    item.click_input()
    time.sleep(0.3)
    _invocar(pasta_dlg, title="OK", control_type="Button")

    fim = time.time() + TIMEOUT_EXPORTACAO
    novo = None
    while time.time() < fim:
        atuais = {p.name for p in PASTA_DOWNLOADS.glob("Fechamento Escrita Fiscal*.xlsx")}
        diferenca = atuais - antes
        if diferenca:
            novo = PASTA_DOWNLOADS / sorted(diferenca)[-1]
            break
        time.sleep(1)
    if novo is None:
        raise RuntimeError(f"Exportação da Planilha de Mercados não gerou arquivo em {TIMEOUT_EXPORTACAO}s.")

    if caminho_destino.exists():
        caminho_destino.unlink()
    novo.replace(caminho_destino)


def _processar_resumo(df_radar: pd.DataFrame, df_mercados: pd.DataFrame, log=None) -> pd.DataFrame:
    """Reproduz a lógica da consulta Power Query do arquivo de referência do
    usuário (Resumo Radar Fiscal.xlsx): limpa Status/RegimeApuracao, junta
    (LEFT JOIN) com a Planilha de Mercados por IdCorporativo=COD trazendo só
    as 6 colunas usadas, e calcula a coluna 'Documentação'."""
    _log("Processando resumo (limpeza + junção)...", log)
    df = df_radar.copy()

    df["Status"] = df["Status"].where(df["Status"].notna(), "Não importado")
    df["Status"] = df["Status"].replace("", "Não Importado")
    df["Status"] = df["Status"].replace(MAPA_STATUS)
    df["RegimeApuracao"] = df["RegimeApuracao"].replace(MAPA_REGIME)
    df["Gerente de Contas"] = df["Gerente de Contas"].apply(
        lambda v: MAPA_GERENTES.get(v.strip().upper(), v) if isinstance(v, str) else v
    )

    mercados = df_mercados.rename(columns={"COD": "IdCorporativo"})
    mercados = mercados[["IdCorporativo"] + COLUNAS_MERCADOS_TRAZIDAS].copy()
    mercados = mercados.rename(columns={c: f"Planilha de Mercados.{c}" for c in COLUNAS_MERCADOS_TRAZIDAS})
    mercados = mercados.dropna(subset=["IdCorporativo"])
    mercados["IdCorporativo"] = mercados["IdCorporativo"].astype(df["IdCorporativo"].dtype)

    resumo = df.merge(mercados, on="IdCorporativo", how="left")

    coluna_ia = "Planilha de Mercados.ÍA"
    resumo["Documentação"] = resumo[coluna_ia].apply(
        lambda v: "Documentação Pendente" if (pd.isna(v) or v == "" or v == "A") else "Documentação Recebida"
    )

    # Dias desde o último comentário da Operação — usado pra sinalizar
    # pendências "paradas" há muito tempo. Fica em branco quando não há
    # comentário registrado (não dá pra inferir "há quanto tempo" sem data).
    coluna_data_op = "Planilha de Mercados.Data Comentário Operação"
    resumo["DiasSemAtualizacao"] = (pd.Timestamp.now() - resumo[coluna_data_op]).dt.days

    ordem = list(df_radar.columns) + [
        coluna_ia, "Documentação", "DiasSemAtualizacao",
        "Planilha de Mercados.DESC. REMESSAS",
        coluna_data_op,
        "Planilha de Mercados.EFV ATUAL Data",
        "Planilha de Mercados.AÇÃO GERENTE",
        "Planilha de Mercados.Data Comentário Gerência",
    ]
    return resumo[ordem]


def _fechar_sistema_analise(log=None):
    _log("Fechando Sistema de Analise...", log)
    _fechar_janela_se_existir("Controle de Análise Fechamento Escrita Fiscal")
    for titulo in ("Radar", "Departamento Fiscal", "Menu"):
        _fechar_janela_se_existir(titulo, exato=(titulo != "Departamento Fiscal"))


def _log(msg, log=None):
    if log:
        log(msg)


def _gerar_json_portal(df: pd.DataFrame) -> None:
    """Gera um JSON enxuto (subconjunto de colunas) para o portal web ler via
    fetch — o portal não faz parsing de .xlsx no navegador."""
    subset = df[COLUNAS_PORTAL].copy()
    subset = subset.rename(columns={
        "Gerente de Contas": "GerenteContas",
        "Planilha de Mercados.AÇÃO GERENTE": "AcaoGerente",
        "Planilha de Mercados.DESC. REMESSAS": "DescRemessas",
    })
    registros = subset.to_dict(orient="records")

    def _serializar(valor):
        # pandas usa NaN/NaT para valores ausentes mesmo em colunas de texto —
        # nenhum dos dois é JSON válido (JS trava no fetch), então vira None
        # (null) aqui. Timestamp (ex.: DataConfirmacao) também não é
        # serializável direto — vira string ISO 8601.
        if pd.isna(valor):
            return None
        if isinstance(valor, pd.Timestamp):
            return valor.isoformat()
        return valor

    registros = [{chave: _serializar(valor) for chave, valor in reg.items()} for reg in registros]
    ARQUIVO_DADOS_PORTAL.write_text(
        json.dumps(registros, ensure_ascii=False, indent=None),
        encoding="utf-8",
    )


# ── Ponto de entrada ─────────────────────────────────────────────────────────

def executar(log=None):
    mgapps = _abrir_mgapps(log)
    # Tudo que abre uma tela do "Sistema de Analise" fica dentro deste
    # try/finally — se qualquer passo falhar no meio (ex.: o menu "Arquivo"
    # não abrir a tempo), o finally garante que Menu/Departamento Fiscal/
    # Radar/Controle de Análise Fechamento Escrita Fiscal são fechados mesmo
    # assim, em vez de ficarem abertos travando a próxima execução (bug real
    # visto ao vivo em 2026-08-21, rodando em sequência com a Análise de
    # Balanço via o orquestrador do Relatório de Fechamentos — mesmo
    # raciocínio que já protegia "Análise de Balanço" lá com seu próprio
    # try/finally).
    try:
        _abrir_sistema_analise(mgapps, log)
        _login_sistema_analise(log)
        fiscal = _abrir_importacao_fiscal(log)

        radar = _abrir_radar_fiscal(fiscal, log)
        _filtrar(radar, log)
        _exportar_excel(radar, ARQUIVO_SAIDA, log)
        _copiar_para_rede(ARQUIVO_SAIDA, ARQUIVO_REDE_RADAR, log)
        _fechar_janela_se_existir("Radar")

        mercados = _abrir_controle_status(fiscal, log)
        _limpar_filtro_mercados(mercados, log)
        _marcar_unidades(mercados, UNIDADES_MERCADOS, log)
        _filtrar_mercados(mercados, log)
        _exportar_planilha_mercados(mercados, ARQUIVO_MERCADOS, log)
        _copiar_para_rede(ARQUIVO_MERCADOS, ARQUIVO_REDE_MERCADOS, log)
    finally:
        _fechar_sistema_analise(log)

    df_radar = pd.read_excel(ARQUIVO_SAIDA)
    df_mercados = pd.read_excel(ARQUIVO_MERCADOS, sheet_name="Gerencial")
    _log(f"Radar Fiscal: {len(df_radar)} registros. Planilha de Mercados: {len(df_mercados)} registros.", log)

    df = _processar_resumo(df_radar, df_mercados, log)
    with pd.ExcelWriter(ARQUIVO_RESUMO, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Resumo", index=False)
    _gerar_json_portal(df)

    ARQUIVO_STATUS.write_text(
        json.dumps(
            {
                "ultima_execucao": datetime.now().isoformat(timespec="seconds"),
                "registros": len(df),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return df


if __name__ == "__main__":
    executar(log=print)
