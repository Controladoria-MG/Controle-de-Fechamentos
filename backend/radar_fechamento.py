"""Robô 1 — Radar de Fechamento (MGApps > Analise Balanço, pywinauto).

Calibrado ao vivo em 2026-08-17 contra o app real. Fluxo:
MG Apps > tile "Analise Balanço" > ícone de pessoa (bonequinho, painel
lateral) > preenche Mês Inicial/Mês Final/Ano (mês anterior) > Pesquisar
> menu "Arquivo" > "Exportar" > escolhe a pasta de destino (diálogo
"Selecionar pasta" moderno, aceita Ctrl+L + caminho digitado).

Achados importantes dessa calibração:
- O app "Analise Balanço" é um .exe separado do MGApps (não uma tela
  dentro do "Sistema de Analise"). Abre sem pedir login separado (usa a
  sessão já autenticada do Windows/MGApps) — MGAPPS_USUARIO/SENHA no
  .env não são usados aqui, mas ficam guardados caso isso mude.
- O botão "Exportar" (dentro do dropdown do menu "Arquivo") não tem
  nome acessível via UI Automation (ícone+texto parecem ser uma imagem
  única) — é clicado por coordenada relativa à janela (128, 172), que
  se mostrou estável independente da posição da janela na tela.
- O diálogo de destino ("Selecionar pasta") é o seletor moderno do
  Windows (não o antigo SHBrowseForFolder) — aceita Ctrl+L pra focar a
  barra de endereço e digitar um caminho direto, então dá pra apontar
  direto pra pasta do projeto (diferente do Radar Fiscal, que precisa
  escolher "Downloads" e depois mover o arquivo).
- O nome do arquivo gerado não é controlado por nós (o diálogo só
  escolhe a pasta) — vem como "Radar MM-AAAA DD-MM-AAAA HH-MM-SS.xlsx".
  O robô localiza o arquivo novo comparando o conteúdo da pasta antes/
  depois do clique em "Selecionar pasta", igual à Planilha de Mercados
  do Radar Fiscal.
"""

import ctypes
import subprocess
import time
from datetime import date, timedelta
from pathlib import Path

import win32api
import win32clipboard
import win32con
import win32gui
import win32process
from dotenv import load_dotenv
from pywinauto import Application
from pywinauto.keyboard import send_keys

RAIZ = Path(__file__).parent.parent
load_dotenv(RAIZ / ".env")

MGAPPS_EXE = r"C:\Program Files (x86)\MGApps\MGApps.Presentation.WpfApp.exe"
PASTA_DESTINO = RAIZ / "data" / "analise_balanco"
ARQUIVO_SAIDA = PASTA_DESTINO / "radar_fechamento.xlsx"

TIMEOUT_ELEMENTO = 20
TIMEOUT_EXPORTACAO = 90


def _log(msg, log=None):
    if log:
        log(msg)


def _garantir_estacao_desbloqueada():
    """A estação pode ser bloqueada (Win+L) durante a execução — sem essa
    checagem, automação de mouse (SetCursorPos) trava com um erro genérico
    do Windows em vez de falhar com uma mensagem clara (visto na calibração
    ao vivo de 2026-08-17)."""
    desktop = ctypes.windll.user32.OpenInputDesktop(0, False, 0)
    if not desktop:
        raise RuntimeError(
            "A estação Windows parece estar bloqueada (Win+L) ou sem sessão "
            "interativa — automação de mouse/teclado não funciona nesse estado. "
            "Desbloqueie a tela e rode de novo."
        )
    ctypes.windll.user32.CloseDesktop(desktop)


# ── Utilitários de janela (mesmo padrão do Radar-Fiscal/backend/radar_fiscal.py) ──

def _janela_por_titulo(titulo, exato=True, timeout=TIMEOUT_ELEMENTO):
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


# ── Rodar em segundo plano (mesma técnica de radar_fiscal.py) ────────────────
#
# A janela "Análise de Balanço" é WPF e responde a UI Automation (.invoke() /
# .select()) sem precisar estar visível ou em primeiro plano — então, depois
# de aberta, ela é "estacionada" fora da tela (_ocultar_janela) e o filtro de
# competência roda sem tomar o mouse/teclado do usuário. As exceções (que
# exigem a janela na tela por ~2s) são: o tile "Analise Balanço" do launcher
# MG Apps (WPF, só responde a clique de mouse real) e o menu "Arquivo" >
# "Exportar" (o flyout WPF não processa PostMessage de forma confiável —
# 0/3 num teste). O diálogo nativo "Selecionar pasta" fica visível o tempo
# que leva pra apontar o destino (rápido, sem digitação — ver _colar_no_campo).

def _ocultar_janela(hwnd: int):
    """'Estaciona' a janela bem longe de qualquer monitor (não minimiza).
    .invoke()/.select()/PostMessage funcionam do mesmo jeito com a janela
    fora da tela — isso só tira ela da vista do usuário."""
    win32gui.SetWindowPos(
        hwnd, 0, -32000, -32000, 0, 0,
        win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE,
    )


def _restaurar_janela(hwnd: int):
    """Traz a janela de volta pra uma posição visível — usada só pra os ~2s
    de clique real no menu 'Arquivo' > 'Exportar', antes de escondê-la de novo."""
    win32gui.SetWindowPos(
        hwnd, 0, 80, 80, 0, 0,
        win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE,
    )


def _invocar(janela, **criterios):
    """Aciona um botão/item de menu via UI Automation (InvokePattern), sem
    mover o mouse nem precisar que a janela esteja em primeiro plano."""
    janela.child_window(**criterios).invoke()


def _tentar_primeiro_plano(hwnd: int) -> bool:
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
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass  # "No error message is available" — Windows às vezes recusa o foco sem motivo; o retry externo cobre isso
    finally:
        if anexado:
            win32process.AttachThreadInput(current_thread, fg_thread, False)
    return win32gui.GetForegroundWindow() == hwnd


def _forcar_primeiro_plano(hwnd: int, tentativas: int = 5, log=None) -> bool:
    """SetForegroundWindow falha intermitentemente no Windows mesmo com
    AttachThreadInput (visto na calibração ao vivo de 2026-08-17) — sem
    erro determinístico, então a mitigação é retry com backoff."""
    espera = 0.4
    for tentativa in range(1, tentativas + 1):
        if _tentar_primeiro_plano(hwnd):
            return True
        _log(f"Foco da janela recusado pelo Windows (tentativa {tentativa}/{tentativas}), tentando de novo...", log)
        time.sleep(espera)
        espera = min(espera * 1.5, 3.0)
    return win32gui.GetForegroundWindow() == hwnd


def _clicar_com_seguranca(elemento, hwnd_janela: int, log=None):
    """Clique real (mouse) — usado nos tiles do launcher MG Apps (WPF), que
    não respondem a UI Automation. Confirma foco + que o ponto de clique
    pertence à janela certa antes de clicar, igual ao Radar Fiscal."""
    if not _forcar_primeiro_plano(hwnd_janela, log=log):
        raise RuntimeError("Não foi possível trazer a janela do MGApps para primeiro plano.")
    rect = elemento.rectangle()
    x, y = rect.mid_point()
    hwnd_no_ponto = win32gui.WindowFromPoint((x, y))
    _, pid_no_ponto = win32process.GetWindowThreadProcessId(hwnd_no_ponto)
    _, pid_janela = win32process.GetWindowThreadProcessId(hwnd_janela)
    if pid_no_ponto != pid_janela:
        raise RuntimeError("O ponto de clique não pertence à janela do MGApps.")
    elemento.click_input()


def _competencia_mes_anterior() -> tuple[int, int]:
    hoje = date.today()
    primeiro_dia_mes_atual = hoje.replace(day=1)
    ultimo_dia_mes_anterior = primeiro_dia_mes_atual - timedelta(days=1)
    return ultimo_dia_mes_anterior.month, ultimo_dia_mes_anterior.year


def _matar_processo_da_janela(hwnd: int, log=None):
    """Encerra à força o processo dono da janela. `.close()` via UI
    Automation não fecha esses apps de fato — visto ao vivo em 2026-08-24:
    a janela 'Análise de Balanço' (e o MG Apps) continuavam abertos, só
    'estacionados' fora da tela, mesmo com o `.close()` rodando sem erro
    dentro do try/finally. Como esses processos existem só pra essa
    automação (a próxima execução sempre fecha e reabre do zero, nunca
    reaproveita), matar por PID é seguro."""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        handle_proc = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(handle_proc, 0)
        win32api.CloseHandle(handle_proc)
    except Exception as e:
        _log(f"AVISO: não consegui encerrar o processo da janela ({e}).", log)


def _fechar_analise_balanco_se_existir(log=None):
    """Fecha a janela 'Análise de Balanço' se já estiver aberta (execução
    anterior que não chegou a fechar, crash, ou o usuário abriu manualmente)
    — a extração sempre começa do zero em vez de reaproveitar uma janela com
    estado desconhecido (filtro parcialmente preenchido, grade com resultado
    de outra competência, etc.), que foi a origem de mais de um bug visto na
    calibração ao vivo de 2026-08-18."""
    try:
        handle = _janela_por_titulo("Análise de Balanço", timeout=2)
    except TimeoutError:
        return
    _log("'Análise de Balanço' já estava aberto — fechando para começar do zero...", log)
    _matar_processo_da_janela(handle, log)
    time.sleep(1)


# ── Passos do fluxo ──────────────────────────────────────────────────────────

def _fechar_mgapps_se_existir(log=None):
    """Fecha a janela 'MG Apps' se já estiver aberta (execução anterior que
    não chegou a fechar, crash, ou o usuário abriu manualmente) — a
    automação sempre começa do zero em vez de reaproveitar uma janela com
    estado desconhecido, mesmo raciocínio de _fechar_analise_balanco_se_existir()
    acima, agora aplicado também ao launcher (2026-08-21: reaproveitar um
    MGApps 'preso' de uma execução anterior que crashou causou falhas em
    cascata ao rodar em sequência com o Radar Fiscal via o orquestrador do
    Relatório de Fechamentos)."""
    try:
        handle = _janela_por_titulo("MG Apps", timeout=2)
    except TimeoutError:
        return
    _log("MGApps já estava aberto — fechando para começar do zero...", log)
    _matar_processo_da_janela(handle, log)
    time.sleep(1)


def _abrir_mgapps(log=None):
    _fechar_mgapps_se_existir(log)
    _log("Abrindo MGApps...", log)
    subprocess.Popen([MGAPPS_EXE])
    handle = _janela_por_titulo("MG Apps", timeout=20)
    mgapps = _conectar(handle)
    mgapps.restore()
    _forcar_primeiro_plano(handle, log=log)
    time.sleep(1)
    return mgapps


def _garantir_tela_sistemas(mgapps, log=None):
    contatos = mgapps.child_window(title="Contatos", control_type="Text")
    if contatos.exists():
        _log("MGApps estava na tela 'Contatos', voltando para 'Sistemas'...", log)
        voltar = mgapps.child_window(auto_id="PART_Header", control_type="Custom").children(control_type="Button")[0]
        voltar.click_input()
        time.sleep(1)


def _abrir_analise_balanco(mgapps, log=None):
    _garantir_tela_sistemas(mgapps, log)
    _log("Procurando tile 'Analise Balanço'...", log)
    tile = mgapps.child_window(title="Analise Balanço", control_type="Text")
    tile.wait("exists", timeout=TIMEOUT_ELEMENTO)
    time.sleep(0.5)

    for tentativa in range(1, 4):
        _clicar_com_seguranca(tile, mgapps.handle, log)
        try:
            handle = _janela_por_titulo("Análise de Balanço", timeout=8)
            _ocultar_janela(handle)  # daqui pra frente roda em segundo plano
            return _conectar(handle)
        except TimeoutError:
            _log(f"Clique não abriu 'Analise Balanço' (tentativa {tentativa}/3), tentando de novo...", log)
    raise TimeoutError("Não foi possível abrir 'Analise Balanço' após 3 tentativas.")


def _abrir_radar(win, log=None):
    _log("Abrindo tela do Radar (ícone de pessoa)...", log)
    listbox = win.child_window(control_type="List")
    itens = listbox.children(control_type="ListItem")
    bonequinho = itens[1].children(control_type="Button")[0]
    bonequinho.invoke()
    time.sleep(1.5)

    radar = win.child_window(auto_id="Radar", control_type="Custom")
    radar.wait("exists", timeout=TIMEOUT_ELEMENTO)
    return radar


def _texto_combo(combo) -> str:
    """O Edit interno do combo fica sempre vazio nesse app (visto ao vivo em
    2026-08-18) mesmo com a seleção funcionando de verdade — `selected_text()`
    é quem reflete o valor selecionado de fato."""
    try:
        texto = combo.selected_text()
        if texto:
            return texto.strip()
    except Exception:
        pass
    edits = combo.children(control_type="Edit")
    return edits[0].window_text().strip() if edits else ""


def _selecionar_combo(combo, valor: str, log=None):
    for tentativa in range(1, 4):
        combo.select(valor)
        time.sleep(0.6)
        atual = _texto_combo(combo)
        if atual == valor:
            return
        _log(f"Combo mostrou {atual!r} em vez de {valor!r} (tentativa {tentativa}/3), tentando de novo...", log)
        time.sleep(0.4)
    raise RuntimeError(f"Não consegui selecionar {valor!r} no combo (ficou em {atual!r}).")


def _filtrar_mes_anterior(win, radar, log=None):
    mes, ano = _competencia_mes_anterior()
    _log(f"Filtrando competência {mes:02d}/{ano}...", log)

    groupbox = radar.child_window(control_type="Group")
    filhos = groupbox.children()
    combo_mes_ini, combo_mes_fim, combo_ano = filhos[0], filhos[1], filhos[2]
    btn_pesquisar = filhos[5]

    _selecionar_combo(combo_mes_ini, str(mes), log)
    _selecionar_combo(combo_mes_fim, str(mes), log)
    _selecionar_combo(combo_ano, str(ano), log)

    btn_pesquisar.invoke()
    time.sleep(4)
    return mes, ano


def _colar_no_campo(edit, texto: str, log=None):
    """Coloca `texto` no Edit sem digitar caractere por caractere: 1º tenta
    set_edit_text (ValuePattern, instantâneo); se não pegar, cai pra colar
    via área de transferência (Ctrl+V, preservando o clipboard do usuário);
    último recurso, digita de verdade."""
    try:
        edit.set_edit_text(texto)
        time.sleep(0.2)
        if (edit.get_value() or "").strip().rstrip("\\") == texto.rstrip("\\"):
            return
    except Exception:
        pass

    guardado = None
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                guardado = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, texto)
        finally:
            win32clipboard.CloseClipboard()
        edit.set_focus()
        send_keys("^a")
        send_keys("^v")
        time.sleep(0.3)
    except Exception:
        _log("Colar falhou — digitando o caminho...", log)
        edit.set_focus()
        send_keys("^a")
        send_keys(texto.replace(" ", "{SPACE}"))
    finally:
        if guardado is not None:
            try:
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, guardado)
                win32clipboard.CloseClipboard()
            except Exception:
                pass


def _exportar(win, radar, caminho_destino: Path, log=None):
    """Abre o menu 'Arquivo' > 'Exportar' e aponta o destino no diálogo
    'Selecionar pasta'.

    O menu 'Arquivo' e o item 'Exportar' (botão sem nome acessível via UIA,
    ícone+texto numa imagem só) são clicados de verdade — a via por mensagem
    (PostMessage) se mostrou instável no flyout WPF (0/3 num teste de
    2026-08-31), então a janela é trazida pra tela só pra esses ~2s de
    clique e escondida de novo logo depois. Cada tentativa fecha um dropdown
    pendente com ESC e reabre do zero; a confirmação é o diálogo aparecendo.

    O caminho no diálogo é preenchido no campo 'Pasta:' (auto_id 1152) sem
    digitação (set_edit_text / colar), pra não gastar segundos digitando
    caractere a caractere."""
    _log("Abrindo menu Arquivo > Exportar...", log)
    _garantir_estacao_desbloqueada()
    _restaurar_janela(win.handle)
    win.set_focus()
    time.sleep(0.3)

    menu_item = radar.child_window(title="Arquivo", control_type="MenuItem")

    dlg_handle = None
    for tentativa in range(1, 4):
        send_keys("{ESC}")  # fecha qualquer dropdown que já esteja aberto
        time.sleep(0.5)
        menu_item.click_input()
        time.sleep(1)
        win.click_input(coords=(128, 172))
        try:
            dlg_handle = _janela_por_titulo("Selecionar pasta", timeout=6)
            break
        except TimeoutError:
            _log(f"Diálogo 'Selecionar pasta' não abriu (tentativa {tentativa}/3), tentando de novo...", log)
    if dlg_handle is None:
        raise RuntimeError("Não consegui abrir o diálogo 'Selecionar pasta' após 3 tentativas.")
    _ocultar_janela(win.handle)  # o Radar já não é mais preciso na tela
    # O diálogo 'Selecionar pasta' é nativo do Windows — mover com SetWindowPos
    # pode travar a conclusão (mesma pegadinha do 'Salvar como' no Radar
    # Fiscal), então ele fica visível esses poucos segundos, não é estacionado.
    dlg = _conectar(dlg_handle)

    caminho_destino.mkdir(parents=True, exist_ok=True)
    antes = {p.name for p in caminho_destino.glob("*.xlsx")}

    _log(f"Apontando destino para {caminho_destino}...", log)
    campo_pasta = dlg.child_window(auto_id="1152", control_type="Edit")
    _colar_no_campo(campo_pasta, str(caminho_destino), log)
    time.sleep(0.5)

    try:
        _invocar(dlg, auto_id="1", control_type="Button")  # botão "Selecionar pasta"
    except Exception:
        dlg.child_window(title="Selecionar pasta", control_type="Button").click_input()

    fim = time.time() + TIMEOUT_EXPORTACAO
    novo = None
    while time.time() < fim:
        atuais = {p.name for p in caminho_destino.glob("*.xlsx")}
        diff = atuais - antes
        if diff:
            novo = caminho_destino / sorted(diff)[-1]
            break
        time.sleep(1)
    if novo is None:
        raise RuntimeError(f"Exportação não gerou arquivo em {TIMEOUT_EXPORTACAO}s.")
    return novo


def executar(log=None) -> Path:
    _garantir_estacao_desbloqueada()
    _fechar_analise_balanco_se_existir(log)
    mgapps = _abrir_mgapps(log)
    win = _abrir_analise_balanco(mgapps, log)
    try:
        radar = _abrir_radar(win, log)
        _filtrar_mes_anterior(win, radar, log)
        novo = _exportar(win, radar, PASTA_DESTINO, log)

        if ARQUIVO_SAIDA.exists():
            ARQUIVO_SAIDA.unlink()
        novo.replace(ARQUIVO_SAIDA)
        _log(f"Radar de Fechamento salvo em {ARQUIVO_SAIDA}", log)
        return ARQUIVO_SAIDA
    finally:
        _log("Fechando 'Análise de Balanço'...", log)
        _matar_processo_da_janela(win.handle, log)


if __name__ == "__main__":
    executar(log=print)
