"""Robô 2 — Retorno do Checklist (Intranet, selenium).

Automatiza https://aplicativo.mgcontecnica.com.br/#/home para baixar o
relatório personalizado "(Meu) Retorno do Checklist" filtrado pela
competência do mês anterior. Lógica de login/navegação/download
reaproveitada de "Utilitarios/Utilitarios/3 - Relatório_Personalizado.py"
(fora deste projeto), que já baixa esse mesmo relatório.
"""

import os
import time
from datetime import date, timedelta
from pathlib import Path

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

RAIZ = Path(__file__).parent.parent
load_dotenv(RAIZ / ".env")

URL_LOGIN = "https://aplicativo.mgcontecnica.com.br/#/home"
USUARIO = os.getenv("INTRANET_USUARIO", "")
SENHA = os.getenv("INTRANET_SENHA", "")

PASTA_DESTINO = RAIZ / "data" / "analise_balanco"
# Download vai para uma subpasta temporária, não direto em PASTA_DESTINO —
# essa pasta é compartilhada com o Robô 1 (radar_fechamento.xlsx), e limpar
# PASTA_DESTINO inteira apagava o arquivo que ele acabou de salvar ali
# quando os dois rodam em sequência pelo orquestrador.
PASTA_TEMP = PASTA_DESTINO / "_temp"
ARQUIVO_SAIDA = PASTA_DESTINO / "retorno_checklist.xlsx"

# Texto exato confirmado em teste real (2026-08-17): "(Meu)Retorno do
# Checklist" (sem espaço após o parêntese). Usa normalize-space + igualdade
# exata porque também existe "(Meu)Retorno do Checklist (EF)" no mesmo menu,
# que um contains() genérico pegaria por engano.
XPATH_MENU_RELATORIO = "//h5[normalize-space(text())='(Meu)Retorno do Checklist']"

TIMEOUT_PADRAO = 20
TIMEOUT_DOWNLOAD = 120


def _competencia_mes_anterior() -> str:
    hoje = date.today()
    primeiro_dia_mes_atual = hoje.replace(day=1)
    ultimo_dia_mes_anterior = primeiro_dia_mes_atual - timedelta(days=1)
    return ultimo_dia_mes_anterior.strftime("%m/%Y")


def _log(msg, log=None):
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
    options.add_argument("--safebrowsing-disable-download-protection")

    driver = webdriver.Chrome(options=options)
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": str(pasta_destino)},
    )
    return driver


def _limpar_pasta(pasta: Path):
    pasta.mkdir(parents=True, exist_ok=True)
    for arquivo in os.listdir(pasta):
        caminho = pasta / arquivo
        if caminho.is_file():
            caminho.unlink()


def executar(log=None) -> Path:
    competencia = _competencia_mes_anterior()
    _log(f"Competência (mês anterior): {competencia}", log)

    _limpar_pasta(PASTA_TEMP)

    driver = _criar_driver(PASTA_TEMP)
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)

    try:
        _log("Abrindo Intranet...", log)
        driver.get(URL_LOGIN)
        time.sleep(2)

        # LOGIN
        _log("Fazendo login...", log)
        try:
            campo_usuario = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "usuario"))
            )
            campo_usuario.send_keys(USUARIO)
            driver.find_element(By.ID, "senha").send_keys(SENHA)
            wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Entrar']"))
            ).click()
            time.sleep(2)
        except Exception:
            pass  # já logado

        # MG CONTROLE
        _log("Acessando MG Controle...", log)
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//h6[@title='MG Controle']"))
        ).click()
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[-1])

        # RELATÓRIOS > PERSONALIZADOS > (MEU) RETORNO DO CHECKLIST
        _log("Navegando até Relatórios > Personalizados > (Meu) Retorno do Checklist...", log)
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//h5[contains(text(),'Relatórios')]"))
        ).click()
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//h5[contains(text(),'Personalizados')]"))
        ).click()
        wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_MENU_RELATORIO))
        ).click()

        # COMPETÊNCIA (mês anterior)
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

        # EXPORTAR
        _log("Exportando relatório...", log)
        wait.until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_ExportarRelatorioLinkButton"))
        ).click()

        # AGUARDAR DOWNLOAD
        _log("Aguardando download...", log)
        arquivo_final = None
        tempo_inicio = time.time()
        while (time.time() - tempo_inicio) < TIMEOUT_DOWNLOAD:
            arquivos = os.listdir(PASTA_TEMP)
            if any(arq.endswith(".crdownload") for arq in arquivos):
                time.sleep(2)
                continue
            arquivos_xlsx = [f for f in arquivos if f.endswith(".xlsx")]
            if arquivos_xlsx:
                arquivo_final = arquivos_xlsx[0]
                break
            time.sleep(2)

        if not arquivo_final:
            raise RuntimeError(f"Download não detectado dentro de {TIMEOUT_DOWNLOAD}s.")

        caminho_baixado = PASTA_TEMP / arquivo_final
        if ARQUIVO_SAIDA.exists():
            ARQUIVO_SAIDA.unlink()
        caminho_baixado.replace(ARQUIVO_SAIDA)
        _limpar_pasta(PASTA_TEMP)

        _log(f"Retorno do Checklist salvo em {ARQUIVO_SAIDA}", log)
        return ARQUIVO_SAIDA

    finally:
        time.sleep(2)
        driver.quit()


if __name__ == "__main__":
    executar(log=print)
