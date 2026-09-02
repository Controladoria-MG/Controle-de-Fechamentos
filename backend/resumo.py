"""Junta as duas exportações brutas (Radar de Fechamento + Retorno do
Checklist) num único resumo, aplicando a regra de negócio confirmada pelo
usuário em 2026-08-17 (ver documentacao.txt):

    tarefa do checklist "Em aberto" ou "Atrasada"  -> Documentação Pendente
    tarefa do checklist "Baixada"                  -> Documentação Recebida
    empresa sem nenhuma tarefa de retorno           -> Documentação Pendente

Confirmado em teste real (2026-08-17) contra os dois arquivos baixados de
verdade pelos robôs: o relatório da Intranet vem com 2 abas ("Totalizador"
e "Pendencias") — os dados por tarefa estão em "Pendencias", não na
primeira aba (que é só um resumo agregado, sem ligação com empresa).

Chave de junção: `IdCliente` (Radar de Fechamento) = `CodCliente` (Retorno
do Checklist) — CONFIRMADO comparando os dois arquivos reais baixados na
mesma sessão: 1626 códigos em comum, nome da empresa batendo 100% nos
casos testados. Mais confiável que juntar por nome (`Cliente`/
`RazaoSocial` têm duplicidade — filiais com nomes iguais, códigos
diferentes).

Pode haver mais de uma linha (tarefa) por empresa no Retorno do Checklist.
Nesse caso, uma pendente/atrasada vence sobre uma baixada: a empresa só
fica "Documentação Recebida" se TODAS as tarefas dela estiverem baixadas.

Além da junção, também limpa `Status` e `Tributacao` do Radar de Fechamento
com a MESMA lógica usada no Radar Fiscal (`backend/radar_fiscal.py` do
projeto Radar-Fiscal, `_processar_resumo`/`MAPA_REGIME`): `Status` vazio/nulo
vira "Não Importado"; `Tributacao` tem o prefixo "Federal -" removido e as 3
variantes de Lucro Real ("L Real -Trimestral", "L.Real - Mensal", "Lucro
Real - Anual") colapsadas num único grupo "Lucro Real" — ver
`MAPA_TRIBUTACAO` abaixo.
"""

from pathlib import Path

import pandas as pd

ABA_CHECKLIST = "Pendencias"
COLUNA_CHAVE_RADAR = "IdCliente"
COLUNA_CHAVE_CHECKLIST = "CodCliente"
COLUNA_STATUS_CHECKLIST = "Status"

STATUS_PENDENTE = {"em aberto", "atrasada", "atrasado", "atraso", "em atraso"}
STATUS_RECEBIDA = {"baixada", "baixado"}

DOC_PENDENTE = "Documentação Pendente"
DOC_RECEBIDA = "Documentação Recebida"

STATUS_NAO_IMPORTADO = "Não Importado"

# Mesma lógica de limpeza da coluna de regime tributário usada no Radar
# Fiscal (MAPA_REGIME em backend/radar_fiscal.py do projeto Radar-Fiscal) —
# remove o prefixo "Federal -" e colapsa as 3 variantes de Lucro Real num
# único grupo.
MAPA_TRIBUTACAO = {
    "Federal - Lucro Presumido": "Lucro Presumido",
    "Federal - L Real -Trimestral": "Lucro Real",
    "Federal - SN": "Simples Nacional",
    "Federal - MEI": "MEI",
    "Federal - Lucro Real - Anual": "Lucro Real",
    "Federal - L.Real - Mensal": "Lucro Real",
    "Federal - Imune": "Imune",
}


def _status_empresa(status_valores: pd.Series) -> str:
    """Aplica a regra: se QUALQUER tarefa da empresa estiver pendente/
    atrasada, a empresa é Pendente; só é Recebida se TODAS as tarefas
    estiverem Baixada."""
    normalizados = status_valores.dropna().astype(str).str.strip().str.lower()
    if any(v in STATUS_PENDENTE for v in normalizados):
        return DOC_PENDENTE
    if len(normalizados) > 0 and all(v in STATUS_RECEBIDA for v in normalizados):
        return DOC_RECEBIDA
    return DOC_PENDENTE


def processar_resumo(df_radar: pd.DataFrame, df_checklist: pd.DataFrame) -> pd.DataFrame:
    if COLUNA_CHAVE_RADAR not in df_radar.columns:
        raise KeyError(
            f"Coluna '{COLUNA_CHAVE_RADAR}' não encontrada no Radar de Fechamento. "
            f"Colunas disponíveis: {list(df_radar.columns)}"
        )
    for coluna in (COLUNA_CHAVE_CHECKLIST, COLUNA_STATUS_CHECKLIST):
        if coluna not in df_checklist.columns:
            raise KeyError(
                f"Coluna '{coluna}' não encontrada na aba '{ABA_CHECKLIST}' do Retorno do Checklist. "
                f"Colunas disponíveis: {list(df_checklist.columns)}"
            )

    df_radar = df_radar.copy()
    df_radar["Status"] = df_radar["Status"].where(df_radar["Status"].notna(), STATUS_NAO_IMPORTADO)
    df_radar["Status"] = df_radar["Status"].replace("", STATUS_NAO_IMPORTADO)
    if "Tributacao" in df_radar.columns:
        df_radar["Tributacao"] = df_radar["Tributacao"].replace(MAPA_TRIBUTACAO)

    status_por_empresa = (
        df_checklist.groupby(COLUNA_CHAVE_CHECKLIST)[COLUNA_STATUS_CHECKLIST]
        .apply(_status_empresa)
        .rename("Documentação")
        .rename_axis(COLUNA_CHAVE_RADAR)
    )

    resumo = df_radar.merge(status_por_empresa, on=COLUNA_CHAVE_RADAR, how="left")
    resumo["Documentação"] = resumo["Documentação"].fillna(DOC_PENDENTE)
    return resumo


def gerar_resumo(arquivo_radar: Path, arquivo_checklist: Path, arquivo_saida: Path) -> pd.DataFrame:
    df_radar = pd.read_excel(arquivo_radar)
    df_checklist = pd.read_excel(arquivo_checklist, sheet_name=ABA_CHECKLIST)
    resumo = processar_resumo(df_radar, df_checklist)
    with pd.ExcelWriter(arquivo_saida, engine="xlsxwriter") as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
    return resumo
