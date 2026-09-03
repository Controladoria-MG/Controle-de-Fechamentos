"""Consolida o relatório "Pendências CTB" (Retorno do Checklist Contábil) —
que vem com UMA linha por documento pendente de cada cliente — em uma linha
por cliente, com os motivos concatenados numa string só.

Essa string vira a coluna `DocumentoPendente` dos registros de Análise de
Balanço no portal (junção por `IdCliente`), do mesmo jeito que o Radar
Fiscal preenche essa coluna a partir da Planilha de Mercados.

O relatório em si é extraído à parte (MG Controle) e deve ser deixado em
`data/analise_balanco/checklist_ctb.xlsx` — este módulo só o consolida.
`backend/orquestrador.py` chama `consolidar()` e faz o LEFT JOIN.

Formato do texto concatenado (escolha do usuário, 2026-09-02 — "Tipo
agrupado, descrições juntas"): um `Tipo` de documento por linha; quando o
mesmo `Tipo` aparece com várias `Descricao`, as descrições viram uma lista
na mesma linha. Descrições longas (às vezes a `Descricao` é uma instrução
inteira) são mantidas na íntegra.

    EXTRATO BANCÁRIO: BRADESCO AG:117 C/C:21739-5; ITAÚ AG: 445 C/C:99355
    EXTRATO APLICAÇÃO FINANCEIRA: BB RF MAIS AUTOMÁTICO
    RELATÓRIO DE CONTAS PAGAS
    RELATÓRIO DO CAIXA
"""

import sys
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent

ABA = "Pendências CTB"
CHAVE = "IdCliente"

PASTA_SAIDA = RAIZ / "data" / "analise_balanco"
# Arquivo bruto que o extrator (a definir) precisa deixar aqui.
ARQUIVO_RELATORIO = PASTA_SAIDA / "checklist_ctb.xlsx"
ARQUIVO_CONSOLIDADO = PASTA_SAIDA / "checklist_ctb_consolidado.xlsx"

# Colunas constantes dentro de um mesmo IdCliente — mantidas no consolidado
# (pega o 1º valor), úteis pra conferência no Excel. A junção no portal usa
# só IdCliente -> DocumentoPendente.
COLUNAS_CLIENTE = [
    "IdCliente", "Cliente", "Grupo", "Segmento", "Tributacao",
    "Competencia", "Unidade", "Gerente", "GerenteMaster", "Status",
]

SEP_ITENS = "\n"
SEP_DESCRICOES = "; "


def _norm(texto: str) -> str:
    return " ".join(str(texto).split()).casefold()


def _motivo_do_cliente(sub: pd.DataFrame) -> str:
    """`sub` tem as colunas Tipo/Descricao das pendências de UM cliente."""
    linhas = []
    for tipo, itens in sub.groupby("Tipo", sort=False):
        # Ignora Descrição vazia ou que só repete o próprio Tipo (acontece
        # bastante — nesses casos a linha fica só com o nome do documento).
        descricoes = [
            d for d in dict.fromkeys(itens["Descricao"].dropna().astype(str).str.strip())
            if d and _norm(d) != _norm(tipo)
        ]
        if descricoes:
            linhas.append(f"{tipo}: {SEP_DESCRICOES.join(descricoes)}")
        else:
            linhas.append(str(tipo))
    return SEP_ITENS.join(linhas)


def _ler_relatorio(caminho_relatorio: Path) -> pd.DataFrame:
    """Lê a aba de pendências do relatório. Usa 'Pendências CTB' se existir;
    senão a 1ª aba (o nome exato do export do portal pode variar)."""
    xls = pd.ExcelFile(caminho_relatorio)
    aba = ABA if ABA in xls.sheet_names else xls.sheet_names[0]
    return xls.parse(aba)


def consolidar(caminho_relatorio: Path = ARQUIVO_RELATORIO) -> pd.DataFrame:
    """Lê o relatório e devolve uma linha por IdCliente com a coluna
    `DocumentoPendente` (motivos concatenados) e `QtdPendencias`."""
    df = _ler_relatorio(caminho_relatorio)

    faltando = {CHAVE, "Tipo", "Descricao"} - set(df.columns)
    if faltando:
        raise KeyError(
            f"Colunas {faltando} não encontradas em {caminho_relatorio.name}. "
            f"Colunas disponíveis: {list(df.columns)}"
        )

    df = df.copy()
    df[CHAVE] = pd.to_numeric(df[CHAVE], errors="coerce")
    df = df.dropna(subset=[CHAVE])
    df[CHAVE] = df[CHAVE].astype("int64")
    df["Tipo"] = df["Tipo"].fillna("").astype(str).str.strip()

    # Linhas idênticas (mesmo cliente + tipo + descrição) contam uma vez só.
    df = df.drop_duplicates(subset=[CHAVE, "Tipo", "Descricao"])

    motivos = (
        df.groupby(CHAVE, sort=False)[["Tipo", "Descricao"]]
        .apply(_motivo_do_cliente)
        .rename("DocumentoPendente")
    )
    qtd = df.groupby(CHAVE, sort=False).size().rename("QtdPendencias")

    colunas_cliente = [c for c in COLUNAS_CLIENTE if c in df.columns]
    base = (
        df.drop_duplicates(subset=[CHAVE])[colunas_cliente]
        .set_index(CHAVE)
    )

    return base.join([qtd, motivos]).reset_index()


def gerar_consolidado(
    caminho_relatorio: Path = ARQUIVO_RELATORIO,
    caminho_saida: Path = ARQUIVO_CONSOLIDADO,
) -> pd.DataFrame:
    consolidado = consolidar(caminho_relatorio)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:
        consolidado.to_excel(writer, sheet_name="Consolidado", index=False)
    return consolidado


if __name__ == "__main__":
    origem = Path(sys.argv[1]) if len(sys.argv) > 1 else ARQUIVO_RELATORIO
    if not origem.exists():
        print(f"Relatório não encontrado: {origem}")
        print("uso: python backend/checklist_ctb.py [caminho do relatório Pendências CTB.xlsx]")
        raise SystemExit(1)
    df = gerar_consolidado(origem)
    print(f"{len(df)} clientes consolidados de {origem.name}")
    print(f"-> {ARQUIVO_CONSOLIDADO}")
