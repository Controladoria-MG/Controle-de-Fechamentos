# CLAUDE.md — Regras de Projeto

## O que é este projeto

Portal único que junta o Radar Fiscal e a Análise de Balanço numa página só, com uma aba grande no topo pra alternar entre "Radar Fiscal" e "Análise de Balanço" — mesma ideia do projeto de Análise de Entrega de SPED (ICMS/Contribuições), aplicada aqui a duas fontes com robôs e esquemas de planilha totalmente diferentes (não só valores diferentes de um mesmo campo).

**Projeto autocontido (migrado em 2026-08-21)**: os robôs de automação (Radar Fiscal via MGApps; Radar de Fechamento + Retorno do Checklist + resumo via MGApps/Intranet) moram dentro deste projeto, em `backend/`. Antes viviam em dois repositórios separados (`Radar-Fiscal/` e `Analise-de-Balanco/`) e eram carregados por caminho — a pedido explícito do usuário ("eles não vão existir mais, somente o Relatório de Fechamentos"), o código foi migrado pra cá e os dois repositórios antigos foram **arquivados no GitHub** (somente leitura, código preservado, não apagados — reversível se precisar).

`python backend/orquestrador.py` roda os dois robôs em sequência, copia/gera os JSONs que o portal lê, e ainda junta os dois resumos num **Excel único** (`data/relatorio_fechamentos/resumo.xlsx`, schema normalizado, pedido explícito do usuário — "quero os dados em excel mesmo").

## Estrutura de Diretórios

```
Raiz/
├── index.html
├── .gitignore
├── .env / .env.example      (USUARIO/SENHA do Radar Fiscal + INTRANET_USUARIO/INTRANET_SENHA da Intranet)
├── CLAUDE.md
├── requirements.txt
├── static/
│   ├── script.js
│   ├── style.css   (classes de aba "tipo-relatorio-*")
│   └── logo.png
├── data/
│   ├── radar_fiscal/            (arquivos brutos do robô Radar Fiscal — .xlsx gitignorados, JSON/status versionados)
│   ├── analise_balanco/         (arquivos brutos do pipeline Análise de Balanço — idem)
│   └── relatorio_fechamentos/
│       ├── radar_fiscal_dados.json       (cópia de data/radar_fiscal/, schema bruto — portal normaliza no frontend)
│       ├── status_radar_fiscal.json      (idem, renomeado)
│       ├── analise_balanco_dados.json    (cópia de data/analise_balanco/)
│       ├── status_analise_balanco.json   (idem, renomeado)
│       ├── status.json                   (execução combinada — total das duas fontes)
│       └── resumo.xlsx                   (as duas fontes juntas, schema normalizado — gerado só pelo orquestrador, não versionado)
└── backend/
    ├── radar_fiscal.py       (robô 1 — MGApps "Sistema de Analise", pywinauto)
    ├── radar_fechamento.py   (robô 2a — MGApps "Analise Balanço", pywinauto)
    ├── retorno_checklist.py  (robô 2b — Intranet, selenium)
    ├── resumo.py             (junta radar_fechamento + retorno_checklist)
    └── orquestrador.py       (roda tudo, gera o Excel único e os JSONs do portal)
```

---

## Regras — OBRIGATÓRIO SEGUIR

### HTML / CSS / JS
Mesmas regras dos outros portais MG: único `.html` é `index.html` na raiz; `.css`/`.js` só em `static/`; nada de lógica de backend em `static/`.

### Dados (`data/`)
- **Nunca editar os JSONs nem o Excel à mão.** Todos são gerados por `backend/orquestrador.py`.
- Pra atualizar com dados novos: `python backend/orquestrador.py`. Roda os dois robôs de verdade (MGApps + Intranet) — leva o tempo normal de cada um somado (~3-4 min na calibração), e precisa rodar numa máquina com MGApps instalado e o `.env` deste projeto preenchido (`USUARIO`/`SENHA` pro Radar Fiscal, `INTRANET_USUARIO`/`INTRANET_SENHA` pra Intranet).
- Em `data/radar_fiscal/` e `data/analise_balanco/`: só os `.xlsx` brutos (e `data/analise_balanco/_temp/`) são ignorados — os JSONs/`status.json` de cada um são versionados (são a entrada dos passos seguintes do orquestrador).
- Em `data/relatorio_fechamentos/`: os 4 JSONs (`radar_fiscal_dados.json`, `status_radar_fiscal.json`, `analise_balanco_dados.json`, `status_analise_balanco.json`) + `status.json` combinado são **versionados** (permite GitHub Pages mostrar dados atualizados a cada push). **`resumo.xlsx` não é versionado** — artefato grande e binário, regenerado a cada execução; quem quiser os dados em Excel roda o orquestrador e abre o arquivo local.

### `.env`
Um único arquivo pras duas fontes (antes eram dois `.env` separados, um por projeto):
```
USUARIO=            # login do Radar Fiscal (Sistema de Analise)
SENHA=
MGAPPS_USUARIO=      # não usado hoje (radar_fechamento.py reaproveita sessão já autenticada do MGApps)
MGAPPS_SENHA=
INTRANET_USUARIO=    # login da Intranet (retorno_checklist.py)
INTRANET_SENHA=
```
**Nunca commitar** — está no `.gitignore`. `orquestrador.py` relê esse arquivo e sobrescreve `radar_fiscal.USUARIO`/`SENHA` e `retorno_checklist.USUARIO`/`SENHA` antes de cada execução (`_recarregar_credenciais()`) — defesa contra `load_dotenv(override=False)` quando este orquestrador roda dentro de um processo compartilhado com outros robôs (hub "Atualização de bases").

### MGApps/Análise de Balanço: fechar sempre, em 3 camadas (2026-08-21, calibração ao vivo)
Rodar os dois robôs em sequência expôs bugs de estado que não existiam rodando cada um sozinho: os dois automatizam o mesmo launcher desktop "MG Apps" e originalmente reaproveitavam a janela se já estivesse aberta. Rodando em sequência isso causava falha em cascata — o segundo robô herdava uma janela "quente"/num estado inesperado e travava clicando em tiles. Corrigido em 3 camadas (a pedido explícito do usuário, em mensagens separadas — "fechar antes de rodar o segundo robô", depois "checar no início também se MG Apps e Análise de Balanço estão abertos"):
1. **`_abrir_mgapps()` em `radar_fiscal.py` e `radar_fechamento.py`** — sempre fecha o "MG Apps" existente (`_fechar_mgapps_se_existir()`) e abre um novo, nunca reaproveita. Mesmo raciocínio que `_fechar_analise_balanco_se_existir()` já usava pra não reaproveitar uma janela "Análise de Balanço" com estado desconhecido.
2. **`radar_fiscal.py`: `_fechar_sistema_analise()` agora roda dentro de um `finally`** — antes só rodava se tudo desse certo antes dela; se `_exportar_excel()` (ou qualquer passo) falhasse no meio, "Sistema de Analise" e suas sub-telas (Menu/Departamento Fiscal/Radar/Controle de Análise Fechamento Escrita Fiscal) ficavam abertas pra sempre. Esse era o bug raiz real por trás do pedido do usuário — fechar só o "MG Apps" (launcher) não adianta se as sub-telas continuam abertas como processos/janelas separados. `radar_fechamento.py` já tinha esse padrão certo desde a calibração de 2026-08-18.
3. **`orquestrador.py::executar()` faz 2 checagens explícitas**: (a) **antes de rodar qualquer robô**, fecha MG Apps (`radar_fiscal._fechar_mgapps_se_existir`) e "Análise de Balanço" (`radar_fechamento._fechar_analise_balanco_se_existir`) se alguma tiver sobrado aberta de uma execução anterior; (b) **entre os dois robôs**, fecha o MG Apps de novo antes de abrir a Análise de Balanço.

Também corrigido nessa calibração: `_exportar_excel()` em `radar_fiscal.py` — o clique no menu "Arquivo" às vezes não abria o popup a tempo; ganhou retry com ESC entre tentativas, mesmo padrão já usado em `_exportar()` (Análise de Balanço).

**Validado ao vivo com sucesso em 2026-08-21** (antes da migração pra dentro deste projeto, com os robôs ainda nos projetos-fonte): pipeline completo, ~3min30s, 4190 registros (2076 Radar Fiscal + 2114 Análise de Balanço), Excel gerado, publicado no GitHub. Depois da migração, uma nova tentativa esbarrou no clique único que abre "Sistema de Analise" (exige a janela do MGApps em primeiro plano de verdade) — não reproduziu nenhum dos bugs de estado corrigidos acima (o `try/finally` fechou tudo certo mesmo com erro), e o sintoma bate com uso concorrente da máquina durante o teste (foreground window era outro app, não o MGApps), não com a migração em si.

### Normalização das duas fontes (`static/script.js` e `backend/orquestrador.py`)
Radar Fiscal e Análise de Balanço têm nomes de coluna diferentes pra conceitos equivalentes. No frontend, `normalizarRadarFiscal()`/`normalizarAnaliseBalanco()` traduzem cada fonte pro mesmo esquema comum **uma vez**, no carregamento (`carregarDados()`), pra todo o resto do arquivo (filtros, cards, ranking, evolução, tabela) trabalhar só com os nomes normalizados. No backend, `_normalizar_radar_fiscal()`/`_normalizar_analise_balanco()` em `orquestrador.py` fazem o mesmo mapeamento em pandas, só pra gerar o Excel único (os JSONs do portal continuam no schema bruto de cada fonte — o frontend é quem normaliza):

| Campo comum | Radar Fiscal | Análise de Balanço |
|---|---|---|
| `Id` | `IdCorporativo` | `IdCliente` |
| `Cliente` | `Nome` | `Cliente` |
| `Grupo` | `Grupo` | `Grupo` |
| `Unidade` | `Unidade` | `Unidade` |
| `Segmento` | `Segmento` | `Segmento` |
| `Gerente` | `GerenteContas` (JSON) / `"Gerente de Contas"` (resumo.xlsx) | `Gerente` |
| `Tributacao` | `RegimeApuracao` | `Tributacao` |
| `Status` | `Status` | `Status` (grafia própria) |
| `Documentacao` | `Documentação` | `Documentação` |
| `Departamento` | `DeptoFiscal` | *(não existe)* |
| `DataReferencia` | `DataConfirmacao` | `DataImportacao` |

`Segmento` e `Documentação` já têm o mesmo nome e os mesmos 2 valores nas duas fontes — não precisam de mapeamento. **Atenção**: o Radar Fiscal usa `"Gerente de Contas"` (com espaço) no `resumo.xlsx`, e só vira `GerenteContas` no JSON do portal (rename feito dentro do próprio `radar_fiscal.py`, `_gerar_json_portal`) — `orquestrador.py` lê direto do `resumo.xlsx`/DataFrame retornado por `radar_fiscal.executar()`, então usa o nome com espaço.

### Tipo de Relatório (aba do topo) — igual ao padrão do SPED
A página inteira (chips de Unidade, filtros, cards, ranking, evolução, tabela) sempre mostra só um Tipo de Relatório por vez (`tipoRelatorioAtivo`), escolhido em `#tipo-relatorio-abas`. Trocar de aba (`selecionarTipoRelatorio()`) reseta todos os filtros e repopula cada `<select>` só com os valores daquela fonte (`repopularSelect()`).

Duas diferenças reais entre as fontes que a UI precisa esconder/adaptar ao trocar de aba:
1. **Departamento só existe no Radar Fiscal.** Os grupos `#f-depto-grupo`/`#t-depto-grupo` recebem a classe `.oculto` quando o Tipo ativo é "Análise de Balanço" (não é só deixado vazio — some da tela). A coluna "Departamento" da tabela **fica sempre visível** (mostra "—" nas linhas da Análise de Balanço) — decisão consciente pra não arriscar desalinhar `<colgroup>`/`<td>` escondendo célula por célula (`display:none` num `<td>` isolado quebra a contagem de colunas do `table-layout: fixed`).
2. **A 2ª aba de cards** é "Por Departamento" (Radar Fiscal, aninhado por Tributação, usa `el.depto`) ou "Por Segmento" (Análise de Balanço, aninhado por Tributação, usa o mesmo `el.segmento` do filtro geral). Config em `QUEBRA_CONFIG_POR_TIPO`; o rótulo do botão (`#quebra-aba-segunda`) é atualizado por `selecionarTipoRelatorio()`.
3. **Os 5 valores reais de `Status`** são diferentes nas duas fontes, cada uma com uma categoria sem equivalente na outra (Radar Fiscal: "Bloqueado"; Análise de Balanço: "Importado Contábil") — por isso `STATUS_ORDEM_POR_TIPO` é por Tipo, não uma lista global única. `STATUS_ROTULOS_POR_TIPO` só ajusta a grafia da Análise de Balanço pra bater com o Radar Fiscal onde o significado é o mesmo (`"OK - Com GC"→"Com o GC"`, `"Não Importado"→"Não importado"`), também usada no filtro/rótulo de tabela (`rotuloStatus()`), não só nos cards.

O resto (ranking por Gerente, "Evolução Diária" por `DataReferencia`, tabela, cards "Por Tributação") é **código idêntico** pras duas fontes — só funciona porque os nomes já estão normalizados.

### `backend/orquestrador.py` — fluxo
1. `_recarregar_credenciais()` — relê `.env`, sobrescreve `radar_fiscal.USUARIO`/`SENHA` e `retorno_checklist.USUARIO`/`SENHA`.
2. Checagem de janelas abertas (MG Apps + Análise de Balanço) de uma execução anterior.
3. `radar_fiscal.executar(log)` → `df_radar`. Fecha MG Apps de novo.
4. `radar_fechamento.executar(log)` → `arquivo_radar_fechamento`; `retorno_checklist.executar(log)` → `arquivo_checklist`; `resumo.gerar_resumo(...)` → `df_balanco`. Gera `analise_balanco_dados.json`/`status.json` em `data/analise_balanco/` (réplica do que o antigo `Analise-de-Balanco/backend/orquestrador.py` fazia — este projeto não tem mais aquele orquestrador intermediário).
5. Copia os JSONs/status de `data/radar_fiscal/` e `data/analise_balanco/` pra `data/relatorio_fechamentos/`.
6. Junta os dois (normalizados) num `resumo.xlsx` único.
7. Escreve `status.json` combinado.

### Como rodar
```
cd Relatorio-de-Fechamentos
python backend/orquestrador.py   # roda os 2 robôs + gera Excel único + gera/copia JSONs do portal
python -m http.server 8794
```
Acessar `http://localhost:8794`.

### Integração com o hub "Atualização de bases"
Registrado como `relatorio_fechamentos` em `backend/relatorios/relatorio_fechamentos.py` do hub — wrapper fino que carrega `backend/orquestrador.py` deste projeto via `importlib.util.spec_from_file_location`, com `sys.path.insert(0, CAMINHO_ROBO.parent)` (necessário desde a migração, porque `orquestrador.py` agora faz imports de módulos irmãos — `import radar_fiscal`, `import radar_fechamento` etc. — que só resolvem com a pasta `backend/` deste projeto no `sys.path`). Depois de rodar, copia os JSONs pro clone em `data/relatorio_fechamentos/` do hub, que é o que `git_manager.publicar()` commita.

### Cores — regra fixa
**Nunca verde/âmbar para status.** Só rampa de vermelho MG, igual a todos os outros dashboards MG.
