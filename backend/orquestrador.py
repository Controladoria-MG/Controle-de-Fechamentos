"""Orquestrador do Relatório de Fechamentos: roda os dois robôs-fonte, cada
um no seu projeto separado — [[Radar-Fiscal]] e [[Analise-de-Balanco]],
pastas irmãs desta —, copia os JSONs que cada um já gera pro portal daqui e
junta os dois resumos num Excel único (schema normalizado). Ponto de
entrada: `python backend/orquestrador.py`.

Este projeto não tem robô próprio — só orquestra os dois já existentes,
carregando cada módulo pelo caminho (mesma técnica que o hub "Atualização
de bases" usa em backend/relatorios/radar_fiscal.py e analise_balanco.py).
"""

import importlib.util
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import dotenv_values

RAIZ = Path(__file__).resolve().parent.parent
PASTA_DESTINO = RAIZ / "data" / "relatorio_fechamentos"
ARQUIVO_RESUMO = PASTA_DESTINO / "resumo.xlsx"

CAMINHO_ROBO_RADAR = RAIZ.parent / "Radar-Fiscal" / "backend" / "radar_fiscal.py"
PASTA_DADOS_RADAR = CAMINHO_ROBO_RADAR.parent.parent / "data" / "radar_fiscal"

CAMINHO_ROBO_BALANCO = RAIZ.parent / "Analise-de-Balanco" / "backend" / "orquestrador.py"
PASTA_DADOS_BALANCO = CAMINHO_ROBO_BALANCO.parent.parent / "data" / "analise_balanco"


def _carregar_modulo(nome, caminho):
    # O orquestrador.py da Análise de Balanço importa radar_fechamento/
    # resumo/retorno_checklist por nome nu (módulos irmãos no mesmo
    # backend/) — sem a pasta dele no sys.path, esse import falha com
    # ModuleNotFoundError ao carregar via spec_from_file_location.
    if str(caminho.parent) not in sys.path:
        sys.path.insert(0, str(caminho.parent))
    spec = importlib.util.spec_from_file_location(nome, caminho)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def _copiar(pasta_origem: Path, nome_arquivo: str, nome_destino: str = None) -> None:
    origem = pasta_origem / nome_arquivo
    if not origem.exists():
        return
    shutil.copy2(origem, PASTA_DESTINO / (nome_destino or nome_arquivo))


# As duas fontes têm nomes de coluna diferentes pra conceitos equivalentes —
# mesmo mapeamento de normalizarRadarFiscal()/normalizarAnaliseBalanco() em
# static/script.js, só que aqui em pandas, pra gerar o Excel único. Os
# JSONs que o portal lê continuam no schema bruto de cada fonte (copiados
# de _copiar(), sem passar por aqui) — só o Excel usa o schema normalizado.
def _normalizar_radar_fiscal(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "TipoRelatorio": "Radar Fiscal",
        "Id": df["IdCorporativo"],
        "Cliente": df["Nome"],
        "Grupo": df["Grupo"],
        "Unidade": df["Unidade"],
        "Segmento": df["Segmento"],
        "Gerente": df["Gerente de Contas"],
        "Tributacao": df["RegimeApuracao"],
        "Status": df["Status"],
        "Documentação": df["Documentação"],
        "Departamento": df["DeptoFiscal"],
        "DataReferencia": df["DataConfirmacao"],
    })


def _normalizar_analise_balanco(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "TipoRelatorio": "Análise de Balanço",
        "Id": df["IdCliente"],
        "Cliente": df["Cliente"],
        "Grupo": df["Grupo"],
        "Unidade": df["Unidade"],
        "Segmento": df["Segmento"],
        "Gerente": df["Gerente"],
        "Tributacao": df["Tributacao"],
        "Status": df["Status"],
        "Documentação": df["Documentação"],
        "Departamento": None,
        "DataReferencia": df["DataImportacao"],
    })


def executar(log=None) -> pd.DataFrame:
    def _log(msg):
        print(msg)
        if log:
            log(msg)

    _log("Carregando os robôs-fonte...")
    robo_radar = _carregar_modulo("radar_fiscal_robo", CAMINHO_ROBO_RADAR)
    robo_balanco = _carregar_modulo("analise_balanco_robo", CAMINHO_ROBO_BALANCO)

    # Checagem antes de qualquer coisa: se uma execução anterior travou/
    # crashou no meio, o MG Apps e/ou a janela "Análise de Balanço" podem ter
    # ficado abertos — começar sempre do zero (fechando os dois se existirem)
    # evita herdar um estado desconhecido de qualquer um dos dois numa nova
    # execução (a pedido do usuário, 2026-08-21).
    _log("Checando se MG Apps/Análise de Balanço ficaram abertos de uma execução anterior...")
    robo_radar._fechar_mgapps_se_existir(log)
    robo_balanco.radar_fechamento._fechar_analise_balanco_se_existir(log)

    _log("Rodando robô do Radar Fiscal...")
    # python-dotenv não sobrescreve variáveis de ambiente já setadas
    # (override=False) — se este orquestrador rodar dentro de um processo
    # compartilhado com outros robôs (ex.: hub "Atualização de bases", que
    # importa vários relatórios no mesmo processo Flask), um USUARIO/SENHA
    # genérico de outro .env pode já estar no ambiente e "vencer" antes do
    # load_dotenv() interno do radar_fiscal.py rodar. Por segurança, lemos o
    # .env do Radar Fiscal direto aqui e sobrescrevemos, mesmo padrão do
    # wrapper radar_fiscal.py do hub.
    env_radar = dotenv_values(CAMINHO_ROBO_RADAR.parent.parent / ".env")
    robo_radar.USUARIO = env_radar.get("USUARIO", robo_radar.USUARIO)
    robo_radar.SENHA = env_radar.get("SENHA", robo_radar.SENHA)
    df_radar = robo_radar.executar(log=log)

    # O Radar Fiscal já fecha as telas do "Sistema de Analise" no final do
    # seu próprio executar() (dentro de um try/finally, mesmo se algo falhar
    # no meio), mas deixa o launcher "MG Apps" em si aberto de propósito
    # (pra reaproveitar depois). Rodando os dois robôs em sequência isso
    # causou falha em cascata (2026-08-21): o robô da Análise de Balanço
    # reaproveitava esse MGApps ainda "quente" da primeira execução e
    # travava procurando o tile 'Analise Balanço'. Por isso aqui fechamos
    # explicitamente antes de começar o segundo robô — reaproveita a própria
    # função que o radar_fiscal.py usa pra fazer o mesmo no início do seu
    # executar(), sem duplicar a lógica.
    robo_radar._fechar_mgapps_se_existir(log)

    _log("Rodando robô da Análise de Balanço...")
    env_balanco = dotenv_values(CAMINHO_ROBO_BALANCO.parent.parent / ".env")
    robo_balanco.retorno_checklist.USUARIO = env_balanco.get("INTRANET_USUARIO", robo_balanco.retorno_checklist.USUARIO)
    robo_balanco.retorno_checklist.SENHA = env_balanco.get("INTRANET_SENHA", robo_balanco.retorno_checklist.SENHA)
    df_balanco = robo_balanco.executar(log=log)

    PASTA_DESTINO.mkdir(parents=True, exist_ok=True)

    _log("Copiando dados do portal de cada fonte...")
    _copiar(PASTA_DADOS_RADAR, "radar_fiscal_dados.json")
    _copiar(PASTA_DADOS_RADAR, "status.json", "status_radar_fiscal.json")
    _copiar(PASTA_DADOS_BALANCO, "analise_balanco_dados.json")
    _copiar(PASTA_DADOS_BALANCO, "status.json", "status_analise_balanco.json")

    _log("Juntando os dois resumos num Excel único...")
    df = pd.concat(
        [_normalizar_radar_fiscal(df_radar), _normalizar_analise_balanco(df_balanco)],
        ignore_index=True,
    )
    with pd.ExcelWriter(ARQUIVO_RESUMO, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Resumo", index=False)

    (PASTA_DESTINO / "status.json").write_text(
        json.dumps(
            {"ultima_execucao": datetime.now().isoformat(timespec="seconds"), "registros": len(df)},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    _log(f"Concluído: {len(df_radar)} Radar Fiscal + {len(df_balanco)} Análise de Balanço = {len(df)} registros.")
    _log(f"Excel: {ARQUIVO_RESUMO}")
    return df


if __name__ == "__main__":
    executar(log=print)
