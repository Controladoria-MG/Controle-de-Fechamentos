"""Orquestrador do Relatório de Fechamentos: roda os dois robôs (Radar
Fiscal via MGApps; e o pipeline Radar de Fechamento + Retorno do Checklist
+ resumo, via MGApps/Intranet), junta os dois resumos num Excel único
(schema normalizado) e gera os JSONs que o portal lê via fetch. Ponto de
entrada: `python backend/orquestrador.py`.

Migrado pra dentro deste projeto em 2026-08-21 — antes os robôs viviam em
dois projetos separados (Radar-Fiscal/ e Analise-de-Balanco/) e eram
carregados por caminho via importlib. A pedido do usuário, este projeto
virou o único portal/robô — os outros dois foram arquivados no GitHub
(código preservado lá, só não mais em uso).
"""

import json
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import dotenv_values, load_dotenv

import radar_fechamento
import radar_fiscal
import resumo as resumo_balanco  # nome evita colisão com o "resumo" (df final) deste arquivo
import retorno_checklist

RAIZ = Path(__file__).resolve().parent.parent
load_dotenv(RAIZ / ".env")

PASTA_DESTINO = RAIZ / "data" / "relatorio_fechamentos"
ARQUIVO_RESUMO = PASTA_DESTINO / "resumo.xlsx"

# radar_fiscal.py já gera seus próprios arquivos (inclusive
# radar_fiscal_dados.json/status.json) em data/radar_fiscal/ — mesma pasta
# de sempre, só que agora relativa a este projeto (RAIZ = Path(__file__)
# dentro de cada módulo aponta pra cá depois da migração).
PASTA_DADOS_RADAR = RAIZ / "data" / "radar_fiscal"

# radar_fechamento.py/retorno_checklist.py gravam em data/analise_balanco/
# (mesmo raciocínio acima); mas, diferente do Radar Fiscal, esses dois não
# têm um orquestrador próprio mais — a junção e a geração do JSON do portal
# pra essa fonte (COLUNAS_PORTAL_BALANCO abaixo) são feitas aqui, réplica do
# que existia em Analise-de-Balanco/backend/orquestrador.py.
PASTA_DADOS_BALANCO = RAIZ / "data" / "analise_balanco"
ARQUIVO_RESUMO_BALANCO = PASTA_DADOS_BALANCO / "resumo.xlsx"

COLUNAS_PORTAL_BALANCO = [
    "IdCliente", "Cliente", "Grupo", "Unidade", "Segmento", "Gerente",
    "Tributacao", "Status", "Documentação", "DataImportacao",
]


def _copiar(pasta_origem: Path, nome_arquivo: str, nome_destino: str = None) -> None:
    import shutil
    origem = pasta_origem / nome_arquivo
    if not origem.exists():
        return
    shutil.copy2(origem, PASTA_DESTINO / (nome_destino or nome_arquivo))


def _serializar(valor):
    # pandas usa NaN/NaT para valores ausentes mesmo em colunas de texto —
    # nenhum dos dois é JSON válido (JS trava no fetch), então vira None
    # (null) aqui. Timestamp também não é serializável direto — vira string
    # ISO 8601.
    if pd.isna(valor):
        return None
    if isinstance(valor, pd.Timestamp):
        return valor.isoformat()
    return valor


def _gerar_json_analise_balanco(df: pd.DataFrame, caminho: Path) -> None:
    """Réplica de _gerar_json_portal() que existia em
    Analise-de-Balanco/backend/orquestrador.py — esse projeto não tem mais
    um orquestrador próprio, então essa etapa (junção já feita por
    resumo_balanco.gerar_resumo() antes de chamar esta função) precisa
    acontecer aqui."""
    subset = df[COLUNAS_PORTAL_BALANCO].copy()
    registros = subset.to_dict(orient="records")
    registros = [{chave: _serializar(valor) for chave, valor in reg.items()} for reg in registros]
    caminho.write_text(json.dumps(registros, ensure_ascii=False, indent=None), encoding="utf-8")


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


def _recarregar_credenciais(log):
    # python-dotenv não sobrescreve variáveis de ambiente já setadas
    # (override=False) — se este orquestrador rodar dentro de um processo
    # compartilhado com outros robôs (ex.: hub "Atualização de bases", que
    # importa vários relatórios no mesmo processo Flask), um USUARIO/SENHA
    # genérico de outro relatório pode já estar no ambiente e "vencer" antes
    # do load_dotenv() deste módulo rodar. Por segurança, lemos o .env
    # direto aqui e sobrescrevemos os atributos de cada módulo antes de
    # cada execução real (também evita precisar reiniciar o servidor ao
    # trocar a senha).
    env = dotenv_values(RAIZ / ".env")
    radar_fiscal.USUARIO = env.get("USUARIO", radar_fiscal.USUARIO)
    radar_fiscal.SENHA = env.get("SENHA", radar_fiscal.SENHA)
    retorno_checklist.USUARIO = env.get("INTRANET_USUARIO", retorno_checklist.USUARIO)
    retorno_checklist.SENHA = env.get("INTRANET_SENHA", retorno_checklist.SENHA)


def executar(log=None) -> pd.DataFrame:
    def _log(msg):
        print(msg)
        if log:
            log(msg)

    _recarregar_credenciais(log)

    # Checagem antes de qualquer coisa: se uma execução anterior travou/
    # crashou no meio, o MG Apps e/ou a janela "Análise de Balanço" podem ter
    # ficado abertos — começar sempre do zero (fechando os dois se existirem)
    # evita herdar um estado desconhecido de qualquer um dos dois numa nova
    # execução.
    _log("Checando se MG Apps/Análise de Balanço ficaram abertos de uma execução anterior...")
    radar_fiscal._fechar_mgapps_se_existir(log)
    radar_fechamento._fechar_analise_balanco_se_existir(log)

    _log("Rodando robô do Radar Fiscal...")
    df_radar = radar_fiscal.executar(log=log)

    # O Radar Fiscal já fecha suas próprias telas do "Sistema de Analise" no
    # final do seu executar() (dentro de um try/finally, mesmo se algo
    # falhar no meio), mas deixa o launcher "MG Apps" em si aberto de
    # propósito (pra reaproveitar depois). Fechamos aqui antes do segundo
    # robô abrir de novo — evita reaproveitar um MGApps "quente" que trava
    # procurando o tile 'Analise Balanço'.
    radar_fiscal._fechar_mgapps_se_existir(log)

    _log("Rodando robô da Análise de Balanço...")
    arquivo_radar_fechamento = radar_fechamento.executar(log)
    arquivo_checklist = retorno_checklist.executar(log)
    df_balanco = resumo_balanco.gerar_resumo(arquivo_radar_fechamento, arquivo_checklist, ARQUIVO_RESUMO_BALANCO)
    _log(f"Análise de Balanço: {len(df_balanco)} registros.")

    PASTA_DADOS_BALANCO.mkdir(parents=True, exist_ok=True)
    _gerar_json_analise_balanco(df_balanco, PASTA_DADOS_BALANCO / "analise_balanco_dados.json")
    (PASTA_DADOS_BALANCO / "status.json").write_text(
        json.dumps(
            {"ultima_execucao": datetime.now().isoformat(timespec="seconds"), "registros": len(df_balanco)},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

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
