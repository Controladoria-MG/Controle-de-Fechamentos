"""Extrator do relatório personalizado "(Meu)checklist contabil em aberto"
(Intranet > MG Controle > Relatórios > Personalizados), selenium.

Lista as tarefas do Retorno do Checklist que ainda estão EM ABERTO na
competência. É a condição de exibição do "Documento Pendente" da Análise
de Balanço (regra do usuário, 2026-09-02: o texto do Checklist Contábil só
aparece quando a tarefa de retorno do checklist daquele cliente está em
aberto).

Estrutura copiada de `backend/retorno_checklist.py` — mesmo tipo de
relatório personalizado, no mesmo menu; muda o item ("(Meu)checklist
contabil em aberto") e a conta de login (esse relatório exige a conta
pessoal do usuário — CHECKLIST_ABERTO_USUARIO/SENHA no .env).

STATUS (2026-09-03): VALIDADO ao vivo com a conta `warruda`
(CHECKLIST_ABERTO_USUARIO/SENHA no .env). Competência 08/2026 → 754 tarefas
"Retorno do Check-List" em aberto, 754 clientes únicos. Confirmados contra
o arquivo real:
  - XPATH_MENU_RELATORIO casou o item ("(Meu)checklist contabil em aberto")
  - aba do xlsx: "Pendencias" (existe também "Totalizador")
  - coluna do código do cliente: "CodCliente"
Junção fim-a-fim (_juntar_checklist_ctb em orquestrador.py) contra o
resumo real da Análise de Balanço: 738 dos 754 preenchidos.
"""

import os
import sys
import time
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

RAIZ = Path(__file__).parent.parent
load_dotenv(RAIZ / ".env")

URL_LOGIN = "https://aplicativo.mgcontecnica.com.br/#/home"
# Conta pessoal do usuário (esse relatório é "(Meu)..."). Cai pra INTRANET_*
# se não houver uma dedicada no .env.
USUARIO = os.getenv("CHECKLIST_ABERTO_USUARIO") or os.getenv("INTRANET_USUARIO", "")
SENHA = os.getenv("CHECKLIST_ABERTO_SENHA") or os.getenv("INTRANET_SENHA", "")

PASTA_DESTINO = RAIZ / "data" / "analise_balanco"
PASTA_TEMP = PASTA_DESTINO / "_temp_aberto"
ARQUIVO_SAIDA = PASTA_DESTINO / "checklist_em_aberto.xlsx"

# "(Meu)checklist contabil em aberto" — casa por trecho, minúsculas e sem
# depender do acento exato (o texto real precisa ser confirmado ao vivo).
XPATH_MENU_RELATORIO = (
    "//h5[contains(translate(normalize-space(.),"
    "'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÂÊ','abcdefghijklmnopqrstuvwxyzáéíóúâê'),"
    "'checklist contabil em aberto')]"
)

TIMEOUT_PADRAO = 20
TIMEOUT_DOWNLOAD = 120

# Nome provável da coluna com o código do cliente (o retorno_checklist.xlsx
# usa "CodCliente" na aba "Pendencias" — esse relatório deve ser o mesmo
# template). Ajustar quando o arquivo real for inspecionado.
COLUNAS_CHAVE_POSSIVEIS = ["CodCliente", "IdCliente", "Cod", "Codigo", "CODIGO"]


def clientes_em_aberto(caminho: Path = ARQUIVO_SAIDA) -> set[int]:
    """Códigos de cliente com tarefa de retorno do checklist EM ABERTO na
    competência. Conjunto vazio se o relatório ainda não foi baixado."""
    if not caminho.exists():
        return set()
    xls = pd.ExcelFile(caminho)
    aba = "Pendencias" if "Pendencias" in xls.sheet_names else xls.sheet_names[0]
    df = xls.parse(aba)
    col = next((c for c in COLUNAS_CHAVE_POSSIVEIS if c in df.columns), None)
    if col is None:
        raise KeyError(
            f"Nenhuma coluna de código de cliente em {caminho.name}. "
            f"Colunas: {list(df.columns)}"
        )
    codigos = pd.to_numeric(df[col], errors="coerce").dropna().astype("int64")
    return set(codigos.tolist())


def _competencia_mes_anterior() -> str:
    hoje = date.today()
    ultimo_dia_mes_anterior = hoje.replace(day=1) - timedelta(days=1)
    return ultimo_dia_mes_anterior.strftime("%m/%Y")


def _log(msg, log=None):
    print(msg)
    if log:
        log(msg)


def _criar_driver(pasta_destino: Path) -> webdriver.Chrome:
    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--log-level=3")
    options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    options.add_experimental_option("prefs", {
        "download.default_directory": str(pasta_destino),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,
    })
    driver = webdriver.Chrome(options=options)
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": str(pasta_destino)},
    )
    return driver


def _limpar_pasta(pasta: Path) -> None:
    pasta.mkdir(parents=True, exist_ok=True)
    for arquivo in os.listdir(pasta):
        caminho = pasta / arquivo
        if caminho.is_file():
            caminho.unlink()


def executar(competencia: str | None = None, log=None) -> Path:
    competencia = competencia or _competencia_mes_anterior()
    _log(f"Checklist Contábil em Aberto — competência {competencia}", log)

    _limpar_pasta(PASTA_TEMP)
    driver = _criar_driver(PASTA_TEMP)
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)

    try:
        driver.get(URL_LOGIN)
        time.sleep(2)

        _log("Fazendo login...", log)
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "usuario"))
            ).send_keys(USUARIO)
            driver.find_element(By.ID, "senha").send_keys(SENHA)
            wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Entrar']"))
            ).click()
            time.sleep(2)
        except Exception:
            pass  # já logado

        _log("Acessando MG Controle...", log)
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//h6[@title='MG Controle']"))
        ).click()
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[-1])

        _log("Navegando: Relatórios > Personalizados > (Meu)checklist contabil em aberto...", log)
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//h5[contains(text(),'Relatórios')]"))
        ).click()
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//h5[contains(text(),'Personalizados')]"))
        ).click()
        wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_MENU_RELATORIO))
        ).click()

        _log(f"Preenchendo competência: {competencia}...", log)
        campo_inicio = wait.until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_DataInicioTextBox"))
        )
        campo_fim = wait.until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_DataFimTextBox"))
        )
        campo_inicio.clear()
        campo_fim.clear()
        campo_inicio.send_keys(competencia)
        campo_fim.send_keys(competencia)

        _log("Exportando relatório...", log)
        wait.until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_ExportarRelatorioLinkButton"))
        ).click()

        _log("Aguardando download...", log)
        arquivo_final = None
        inicio = time.time()
        while (time.time() - inicio) < TIMEOUT_DOWNLOAD:
            arquivos = os.listdir(PASTA_TEMP)
            if any(a.endswith(".crdownload") for a in arquivos):
                time.sleep(2)
                continue
            planilhas = [f for f in arquivos if f.lower().endswith((".xlsx", ".xls"))]
            if planilhas:
                arquivo_final = planilhas[0]
                break
            time.sleep(2)

        if not arquivo_final:
            raise RuntimeError(f"Download não detectado dentro de {TIMEOUT_DOWNLOAD}s.")

        baixado = PASTA_TEMP / arquivo_final
        if ARQUIVO_SAIDA.exists():
            ARQUIVO_SAIDA.unlink()
        baixado.replace(ARQUIVO_SAIDA)
        _limpar_pasta(PASTA_TEMP)

        _log(f"Checklist Contábil em Aberto salvo em {ARQUIVO_SAIDA}", log)
        return ARQUIVO_SAIDA

    finally:
        time.sleep(2)
        driver.quit()


if __name__ == "__main__":
    comp = sys.argv[1] if len(sys.argv) > 1 else None
    executar(comp, log=print)
