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

import checklist_ctb
import checklist_ctb_extrator
import checklist_em_aberto
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
    "Tributacao", "Status", "Documentação", "DataImportacao", "DocumentoPendente",
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


def _juntar_checklist_ctb(df_balanco: pd.DataFrame, log_msg) -> pd.DataFrame:
    """Preenche a coluna `DocumentoPendente` dos registros da Análise de
    Balanço a partir do Checklist Contábil (ver backend/checklist_ctb.py),
    por LEFT JOIN em `IdCliente`, e ajusta a coluna `Documentação` pela mesma
    fonte (ver abaixo).

    Regra do usuário (2026-09-02): o texto de `DocumentoPendente` só vale
    pros clientes com tarefa de retorno do checklist EM ABERTO (relatório
    "(Meu)checklist contabil em aberto", ver backend/checklist_em_aberto.py).
    Sem esse relatório não dá pra aplicar a regra — então a coluna fica
    vazia e a `Documentação` não é mexida.

    Regra do usuário (2026-09-03): "não tem como uma tarefa pendente não ter
    documentação" — se o cliente NÃO tem tarefa de retorno do checklist em
    aberto, não há de onde vir uma pendência de documento, então a
    `Documentação` dele é considerada "Documentação Recebida" (a Análise de
    Balanço marcava "Documentação Pendente" só por o Status ser "Não
    Importado", sem lastro num documento faltando de verdade).

    Nenhum dos dois arquivos ausente derruba o pipeline."""
    df_balanco = df_balanco.copy()
    df_balanco["DocumentoPendente"] = pd.NA
    ids = pd.to_numeric(df_balanco["IdCliente"], errors="coerce")

    caminho = checklist_ctb.ARQUIVO_RELATORIO
    if not caminho.exists():
        log_msg(f"Checklist CTB: {caminho.name} não encontrado — DocumentoPendente ficará vazio.")
        return df_balanco

    abertos = checklist_em_aberto.clientes_em_aberto()
    if not abertos:
        log_msg(
            "Checklist em aberto: relatório ausente — sem a regra 'tarefa em aberto', "
            "DocumentoPendente não é preenchido nem a Documentação é ajustada."
        )
        return df_balanco

    tem_tarefa = ids.isin(abertos)

    consolidado = checklist_ctb.gerar_consolidado(caminho)
    mapa = consolidado.set_index("IdCliente")["DocumentoPendente"]
    df_balanco["DocumentoPendente"] = ids.map(mapa).where(tem_tarefa)

    # Sem tarefa de retorno do checklist em aberto → documentação recebida.
    doc_ajustada = int(
        (~tem_tarefa & (df_balanco["Documentação"] != "Documentação Recebida")).sum()
    )
    df_balanco.loc[~tem_tarefa, "Documentação"] = "Documentação Recebida"

    casaram = int(df_balanco["DocumentoPendente"].notna().sum())
    log_msg(
        f"Checklist CTB: {len(consolidado)} cliente(s) no relatório de recebimento, "
        f"{len(abertos)} com tarefa em aberto, {casaram} preenchidos na Análise de Balanço; "
        f"{doc_ajustada} linha(s) sem tarefa em aberto viraram Documentação Recebida."
    )
    return df_balanco


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
        # No Radar Fiscal vem da Planilha de Mercados (já mascarada por ÍA=="A"
        # em radar_fiscal._processar_resumo); na Análise de Balanço vem do
        # Retorno do Checklist Contábil (checklist_ctb.py). Mesmo destino:
        # coluna DocumentoPendente do portal.
        "DocumentoPendente": df.get("Planilha de Mercados.PENDENCIAS FECHAMENTO"),
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
        "DocumentoPendente": df.get("DocumentoPendente"),
    })


# Único Status "de verdade" compatível com Documentação Pendente — pedido do
# usuário (2026-08-25): "todas as empresas [com documentação pendente]
# precisam estar com status Não importado". É por Tipo de Relatório porque a
# grafia da fonte é diferente ("Não importado" no Radar Fiscal, "Não
# Importado" na Análise de Balanço — mesma diferença de STATUS_ORDEM_POR_TIPO
# em static/script.js).
STATUS_NAO_IMPORTADO_POR_TIPO = {
    "Radar Fiscal": "Não importado",
    "Análise de Balanço": "Não Importado",
}


def _corrigir_documentacao_inconsistente(df: pd.DataFrame) -> pd.DataFrame:
    """O usuário só citou o caso "Fechado" como exemplo, mas checando os dados
    reais (resumo.xlsx de 2026-08-24) qualquer Status diferente de "Não
    importado" aparece às vezes com documentação pendente (Fechado: 600
    linhas; Com o GC/Simulando: mais 145) — pra cumprir a regra geral que ele
    pediu ("todas... com status Não importado"), tratamos qualquer uma dessas
    combinações como inconsistência de dados e normalizamos para
    "Documentação Recebida". Mesma regra replicada em
    static/script.js::corrigirDocumentacao() pro portal — lá o schema já
    normalizado usa a coluna "Documentacao" (sem cedilha/til), aqui ainda é
    "Documentação" (schema normalizado deste módulo).

    **Só Radar Fiscal** (2026-09-03): a Análise de Balanço passou a decidir
    a `Documentação` pela tarefa de retorno do checklist em aberto
    (_juntar_checklist_ctb) — lá "Pendente" pode coexistir com qualquer
    Status (a tarefa está aberta), então a regra de Status não vale mais
    pra ela."""
    pendente = df["Documentação"] == "Documentação Pendente"
    status_esperado = df["TipoRelatorio"].map(STATUS_NAO_IMPORTADO_POR_TIPO)
    inconsistente = (
        (df["TipoRelatorio"] == "Radar Fiscal")
        & pendente
        & (df["Status"] != status_esperado)
    )
    df.loc[inconsistente, "Documentação"] = "Documentação Recebida"
    return df


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

    # Tudo abaixo fica num try/finally pra garantir que o launcher "MG Apps"
    # feche no final mesmo se algum passo falhar no meio (ex.: o clique que
    # abre 'Sistema de Analise' falhando após as 3 tentativas) — antes só
    # fechava depois de tudo dar certo, e uma falha deixava o MG Apps preso
    # até a próxima execução (visto ao vivo em 2026-08-24).
    try:
        _log("Rodando robô do Radar Fiscal...")
        df_radar = radar_fiscal.executar(log=log)

        # O Radar Fiscal já fecha suas próprias telas do "Sistema de Analise"
        # no final do seu executar() (dentro de um try/finally, mesmo se algo
        # falhar no meio), mas o launcher "MG Apps" em si continua aberto.
        # Fechamos aqui antes do segundo robô abrir de novo — evita
        # reaproveitar um MGApps "quente" que trava procurando o tile
        # 'Analise Balanço'.
        radar_fiscal._fechar_mgapps_se_existir(log)

        _log("Rodando robô da Análise de Balanço...")
        arquivo_radar_fechamento = radar_fechamento.executar(log)
        arquivo_checklist = retorno_checklist.executar(log)
        df_balanco = resumo_balanco.gerar_resumo(arquivo_radar_fechamento, arquivo_checklist, ARQUIVO_RESUMO_BALANCO)

        # Baixa o "Checklist Contábil > Recebimento" (detalhe do documento
        # pendente de cada cliente). Se falhar — credencial faltando, portal
        # fora do ar — o pipeline segue e a coluna DocumentoPendente fica
        # vazia (não é bloqueante).
        for nome, extrator in (
            ("Checklist Contábil (Recebimento)", checklist_ctb_extrator),
            ("Checklist Contábil em Aberto", checklist_em_aberto),
        ):
            try:
                _log(f"Baixando {nome}...")
                extrator.executar(log=_log)
            except Exception as erro_ext:  # noqa: BLE001 — não pode derrubar o pipeline
                _log(f"{nome}: extração falhou ({erro_ext}). DocumentoPendente pode ficar vazio.")

        df_balanco = _juntar_checklist_ctb(df_balanco, _log)
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
        df = _corrigir_documentacao_inconsistente(df)
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
    finally:
        radar_fiscal._fechar_mgapps_se_existir(log)

    return df


if __name__ == "__main__":
    executar(log=print)
