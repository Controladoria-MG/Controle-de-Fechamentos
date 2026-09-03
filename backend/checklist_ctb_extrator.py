"""Extrator do relatório "Checklist Contábil > Relatórios > Recebimento"
(Intranet, selenium) — baixa uma linha por documento pendente de cada
cliente na competência dada. É consolidado depois por
backend/checklist_ctb.py (1 linha por cliente) e vira a coluna
`DocumentoPendente` dos registros de Análise de Balanço no portal.

Adaptado do `CTB.py` que o usuário deixou na raiz (2026-09-02), com 4
correções pra rodar no pipeline:
  - `--headless=new` (o original tinha `--headless-new`, flag inválida)
  - credenciais do `.env` (eram fixas no código)
  - competência é parâmetro (era um Prompt interativo)
  - baixa pra `data/analise_balanco/` numa subpasta temporária (o original
    baixava numa pasta fixa da rede e apagava tudo que tinha nela)

Estrutura (driver/login/espera de download) copiada de
`backend/retorno_checklist.py`, que já faz o mesmo tipo de automação nesse
mesmo portal.

STATUS (2026-09-02): ✅ validado ao vivo — rodou de primeira com o
fallback `INTRANET_*`, baixou 8.953 linhas (aba "Pendências CTB").
"""

import os
import sys
import time
from datetime import date, timedelta
from pathlib import Path

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import checklist_ctb  # mesma pasta backend/ — reaproveita o caminho de saída

RAIZ = Path(__file__).parent.parent
load_dotenv(RAIZ / ".env")

URL_LOGIN = "https://aplicativo.mgcontecnica.com.br/#/home"
# Conta específica do Checklist Contábil; cai pra INTRANET_* se não houver
# uma dedicada no .env.
USUARIO = os.getenv("CHECKLIST_CTB_USUARIO") or os.getenv("INTRANET_USUARIO", "")
SENHA = os.getenv("CHECKLIST_CTB_SENHA") or os.getenv("INTRANET_SENHA", "")

ARQUIVO_SAIDA = checklist_ctb.ARQUIVO_RELATORIO  # data/analise_balanco/checklist_ctb.xlsx
PASTA_TEMP = ARQUIVO_SAIDA.parent / "_temp_ctb"

TIMEOUT_PADRAO = 20
TIMEOUT_DOWNLOAD = 180

# O app aparece na home como um card com esse trecho no title.
XPATH_CARD_APP = (
    "//div[contains(@class,'card-app') and contains(@title,'aplicativos/Checklist/Contabil')]"
)


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
    """Baixa o relatório da competência dada (padrão: mês anterior, MM/YYYY)
    e devolve o caminho de `checklist_ctb.xlsx`."""
    competencia = competencia or _competencia_mes_anterior()
    _log(f"Checklist Contábil — competência {competencia}", log)

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

        _log("Abrindo o app Checklist Contábil...", log)
        card = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_CARD_APP)))
        driver.execute_script("arguments[0].scrollIntoView(true);", card)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", card)
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[-1])

        _log("Navegando: Relatórios > Recebimento...", log)
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//h2[contains(text(),'Relatórios')]"))
        ).click()
        time.sleep(1)
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//h2[contains(text(),'Recebimento')]"))
        ).click()
        time.sleep(1)

        # Os campos aceitam AAAA-MM (o CTB.py original monta assim a partir de
        # MM/YYYY); início = fim = a competência.
        data_fmt = f"{competencia[3:]}-{competencia[:2]}"
        _log(f"Preenchendo período {data_fmt}...", log)
        wait.until(EC.presence_of_element_located((By.ID, "MainContent_datainicial")))
        driver.execute_script(
            f"document.getElementById('MainContent_datainicial').value = '{data_fmt}';"
        )
        driver.execute_script(
            f"document.getElementById('MainContent_datafinal').value = '{data_fmt}';"
        )
        time.sleep(0.5)

        _log("Gerando Excel...", log)
        wait.until(
            EC.element_to_be_clickable((By.ID, "MainContent_GeraExcelLinkButton"))
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

        _log(f"Checklist Contábil salvo em {ARQUIVO_SAIDA}", log)
        return ARQUIVO_SAIDA

    finally:
        time.sleep(2)
        driver.quit()


if __name__ == "__main__":
    comp = sys.argv[1] if len(sys.argv) > 1 else None
    executar(comp, log=print)
