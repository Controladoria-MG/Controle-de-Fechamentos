# CLAUDE.md — Regras de Projeto

## O que é este projeto

Portal **"Controle de Fechamentos"** (nome de exibição — pasta/repo continuam `Relatorio-de-Fechamentos`) que junta o Radar Fiscal e a Análise de Balanço numa página só, com uma aba grande no topo pra alternar entre "Radar Fiscal" e "Análise de Balanço" — mesma ideia do projeto de Análise de Entrega de SPED (ICMS/Contribuições), aplicada aqui a duas fontes com robôs e esquemas de planilha totalmente diferentes (não só valores diferentes de um mesmo campo).

**Navegação Unidade → Departamento/Segmento (2026-08-25), no mesmo padrão do Portal de Tarefas**: em vez do filtro de Unidade em chips + aba "Por Departamento" que existiam antes, a página agora tem 3 "telas" controladas por `escopo = { unidade, depto }` em `static/script.js`:
1. **Painel de Unidades** (`escopo.unidade === null`): grid de cards, um por Unidade, no mesmo estilo visual dos cards de detalhamento (total + Documentação Recebida/Pendente). Clicar entra na Unidade.
2. **Tela da Unidade** (`escopo.unidade` setado, `escopo.depto === null`): mostra **as 2 visões ao mesmo tempo** — o corpo inteiro do dashboard (Filtros, Por Tributação, Evolução, Ranking, Tabela) já é a visão **Consolidada** da unidade inteira, e logo acima aparece o grid de cards "Por Departamento" (Radar Fiscal) / "Por Segmento" (Análise de Balanço) pra detalhar.
3. **Tela do Departamento/Segmento** (`escopo.depto` também setado): mesmo corpo do dashboard, recortado só pra aquele departamento/segmento dentro da unidade.

Breadcrumb (`#navegacao-breadcrumb`) navega de volta a qualquer nível. Trocar de Tipo de Relatório (aba do topo) sempre volta pro Painel de Unidades. `dadosEscopo` (= `dadosTipo` recortado por `escopo`) é a base de tudo dentro do corpo do dashboard — os filtros de Unidade/Departamento que existiam nos cards "Filtros"/"Filtros da tabela" foram removidos (a navegação já cobre o que eles faziam).

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

### Automação em segundo plano (2026-08-31)
Os 3 robôs rodam sem tomar o mouse/teclado nem mostrar janela na tela, com estas exceções inevitáveis:
- **`retorno_checklist.py`**: Chrome `--headless=new`, invisível.
- **`radar_fiscal.py`**: `.invoke()`/`PostMessage` + `_ocultar_janela()` (estaciona as janelas em -32000,-32000). Piscam só: o launcher "MG Apps" (visível o tempo todo), 1 clique real no tile "Sistema de Analise" (WPF, não responde a UIA), e o diálogo nativo "Salvar como" (poucos segundos — mover ele trava o salvamento).
- **`radar_fechamento.py`**: `_ocultar_janela()` na janela "Análise de Balanço" logo após abrir — o filtro de competência roda escondido. Pra o menu "Arquivo" > "Exportar" a janela volta pra tela por ~2s (`_restaurar_janela` + `click_input` real: o flyout WPF não processa PostMessage de forma confiável — 0/3 num probe) e é escondida de novo assim que o diálogo abre. O diálogo nativo "Selecionar pasta" fica visível o tempo de apontar o destino — o caminho é posto no campo `auto_id="1152"` sem digitar (`_colar_no_campo`: `set_edit_text` → colar via clipboard → digitar). Tile "Analise Balanço" do launcher = clique real, igual ao Radar Fiscal.

Numa máquina só, o usuário não consegue ter a sessão do robô e a dele ao mesmo tempo (Win 11 Pro = 1 sessão interativa). Pra rodar 100% invisível durante o expediente: RDP desconectado numa 2ª máquina/VM, ou agendar fora do horário.

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
| `DocumentoPendente` | `DocumentoPendente` (só frontend, ver abaixo) | *(não existe)* |

`DocumentoPendente` só existe no frontend (não tem coluna equivalente no `resumo.xlsx`/Excel único) — vem da coluna `"PENDENCIAS FECHAMENTO"` da Planilha de Mercados (só SP/GOIAS passam por ela, ver `UNIDADES_MERCADOS`), trazida pro JSON do portal já com esse nome (rename em `_gerar_json_portal`) e passada por `formatarDocumentoPendente()` no frontend (só remove espaço em branco nas pontas). **Regra de negócio, pedida pelo usuário (2026-08-25)**: só faz sentido mostrar essa pendência quando `"ÍA" == "A"` — confirmado com os dados reais de teste: "PENDENCIAS FECHAMENTO" tinha texto em 132 linhas com `ÍA` variado, mas só as 2 linhas com `ÍA == "A"` deviam contar. **É coluna da tabela** (`Documento Pendente`, última coluna) — chegou a ser tirada a pedido do usuário e depois recolocada na mesma sessão (2026-08-25).

**Atenção (2026-08-25)**: a pasta `backend/` (todos os `.py` do robô — `radar_fiscal.py`, `radar_fechamento.py`, `retorno_checklist.py`, `resumo.py`, `orquestrador.py`) foi removida do repositório fora desta conversa (confirmado pelo usuário como intencional) — a máscara `ÍA == "A"` descrita acima **não existe mais no código atual**, só no histórico do git (commit `75f5058` e anteriores). Se o robô/backend voltar a existir (reescrito ou restaurado do histórico), reaplicar essa máscara em `_processar_resumo()`.

`Segmento` e `Documentação` já têm o mesmo nome e os mesmos 2 valores nas duas fontes — não precisam de mapeamento. **Atenção**: o Radar Fiscal usa `"Gerente de Contas"` (com espaço) no `resumo.xlsx`, e só vira `GerenteContas` no JSON do portal (rename feito dentro do próprio `radar_fiscal.py`, `_gerar_json_portal`) — `orquestrador.py` lê direto do `resumo.xlsx`/DataFrame retornado por `radar_fiscal.executar()`, então usa o nome com espaço.

### Tipo de Relatório (aba do topo) — igual ao padrão do SPED
A página inteira (navegação de Unidade/Departamento, filtros, cards, ranking, evolução, tabela) sempre mostra só um Tipo de Relatório por vez (`tipoRelatorioAtivo`), escolhido em `#tipo-relatorio-abas`. Trocar de aba (`selecionarTipoRelatorio()`) volta pro Painel de Unidades, reseta todos os filtros e repopula cada `<select>` só com os valores daquele escopo (`repopularSelect()`).

Diferenças reais entre as fontes que a UI precisa esconder/adaptar ao trocar de aba:
1. **Departamento só existe no Radar Fiscal.** É a "2ª dimensão" de navegação (`QUEBRA_CONFIG_POR_TIPO[tipo].segunda`) — "Por Departamento" no Radar Fiscal, "Por Segmento" na Análise de Balanço (reaproveita a mesma coluna do filtro geral "Segmento"). A coluna "Departamento" da tabela **fica sempre visível** (mostra "—" nas linhas da Análise de Balanço) — decisão consciente pra não arriscar desalinhar `<colgroup>`/`<td>` escondendo célula por célula (`display:none` num `<td>` isolado quebra a contagem de colunas do `table-layout: fixed`).
2. **Os 5 valores reais de `Status`** são diferentes nas duas fontes, cada uma com uma categoria sem equivalente na outra (Radar Fiscal: "Bloqueado"; Análise de Balanço: "Importado Contábil") — por isso `STATUS_ORDEM_POR_TIPO` é por Tipo, não uma lista global única. `STATUS_ROTULOS_POR_TIPO` só ajusta a grafia da Análise de Balanço pra bater com o Radar Fiscal onde o significado é o mesmo (`"OK - Com GC"→"Com o GC"`, `"Não Importado"→"Não importado"`), também usada no filtro/rótulo de tabela (`rotuloStatus()`), não só nos cards. O último item de cada lista em `STATUS_ORDEM_POR_TIPO` é sempre o equivalente a "não importado" — usado por `renderizarDocGrupo()` (ver regra de Documentação abaixo).

O resto (ranking por Gerente, "Evolução Diária" por `DataReferencia`, tabela, cards "Por Tributação") é **código idêntico** pras duas fontes — só funciona porque os nomes já estão normalizados.

### Documentação Pendente só deveria coexistir com Status "Não importado" (2026-08-25)
Pedido do usuário: "todas as empresas [com documentação pendente] precisam estar com um status Não importado" — na prática, checando os dados reais, isso não era verdade (Status "Fechado" era o caso mais comum, mas "Simulando"/"Com o GC" também apareciam com documentação pendente — mais de 700 das 4190 linhas na calibração de 2026-08-24). Corrigido no **tratamento da base**, não só na exibição — qualquer linha com `Documentação == "Documentação Pendente"` e `Status` diferente do "não importado" daquela fonte é normalizada para `"Documentação Recebida"`:
- Frontend: `corrigirDocumentacao()` em `static/script.js`, chamada de dentro de `normalizarRadarFiscal()`/`normalizarAnaliseBalanco()` (afeta a exibição do portal).
- Backend: `_corrigir_documentacao_inconsistente()` em `backend/orquestrador.py`, chamada depois do `pd.concat()` dos dois normalizados (afeta o `resumo.xlsx`).
- Os JSONs brutos (`radar_fiscal_dados.json`/`analise_balanco_dados.json`) **não** são alterados — continuam no schema bruto de cada fonte, como o resto da normalização (o frontend é quem corrige na hora de montar `dados`).

Como consequência, `renderizarDocGrupo()` simplifica o card de "Documentação Pendente" pra mostrar só 1 linha de status (a "não importado" da fonte) em vez das 5 — as outras seriam sempre zero.

### `backend/orquestrador.py` — fluxo
1. `_recarregar_credenciais()` — relê `.env`, sobrescreve `radar_fiscal.USUARIO`/`SENHA` e `retorno_checklist.USUARIO`/`SENHA`.
2. Checagem de janelas abertas (MG Apps + Análise de Balanço) de uma execução anterior.
3. `radar_fiscal.executar(log)` → `df_radar`. Fecha MG Apps de novo.
4. `radar_fechamento.executar(log)` → `arquivo_radar_fechamento`; `retorno_checklist.executar(log)` → `arquivo_checklist`; `resumo.gerar_resumo(...)` → `df_balanco`. Gera `analise_balanco_dados.json`/`status.json` em `data/analise_balanco/` (réplica do que o antigo `Analise-de-Balanco/backend/orquestrador.py` fazia — este projeto não tem mais aquele orquestrador intermediário).
5. Copia os JSONs/status de `data/radar_fiscal/` e `data/analise_balanco/` pra `data/relatorio_fechamentos/`.
6. Junta os dois (normalizados) num `resumo.xlsx` único.
7. Escreve `status.json` combinado.

### Cards totalizadores (placares) medem o FECHAMENTO (2026-09-02)
Os 3 placares do topo (`renderizarPlacares`): **Total de Empresas / Fechamento Concluído / Fechamento Pendente**. "Concluído" = `Status === "Fechado"` (valor literal idêntico nas duas fontes — `fechamentoConcluido()`/`STATUS_CONCLUIDO`); todo o resto (Simulando, Não importado, Com o GC, Bloqueado, Importado Contábil) entra em "Pendente". `Concluído + Pendente = Total`. Antes os cards mediam Documentação Recebida/Pendente — o usuário pediu a troca ("estão contabilizando o tanto de documento recebido, e não o quanto do fechamento já está concluído"). Concluído/Pendente são clicáveis: **trocam a aba da tabela do fim da página** (Concluído/Pendente) e rolam até ela — não setam um filtro de `<select>` ("pendente" é um conjunto de status, não um valor único).

### Abas Pendente / Concluído da tabela do fim (2026-09-02)
`#tabela-abas` dentro de `.tabela-card`, mesma ideia das abas do Portal de Tarefas. `abaTabelaAtiva` ("Pendente"|"Concluído") é um corte adicional aplicado em `aplicarFiltroTabela()` **depois** dos 6 filtros da tabela: Pendente = `!fechamentoConcluido(r)`, Concluído = `fechamentoConcluido(r)`. `definirAbaTabela()` só troca estado+realce (usado no reset ao entrar em escopo novo, junto com a limpeza de filtros — sempre volta pra "Pendente"); `trocarAbaTabela()` troca + re-renderiza. Os cliques nos cards "Por Tributação" continuam filtrando essa tabela normalmente (a aba é um recorte por cima).

### Modal de registros de uma linha de detalhe (2026-09-02, mesmo padrão do portal de SPED)
`#modal-registros` — abre ao clicar numa **linha de detalhe dentro de um card "Por Tributação"** (o bloco de Documentação Recebida/Pendente `.doc-cabecalho`, ou uma linha de Status `.status-linha`) ou numa **linha do ranking por Gerente**. `stopPropagation` nessas linhas pra não disparar também o clique do card (que filtra a tabela do fim). Recebe o subconjunto já recortado pelo contexto do clique (a partir de `filtrados`, que já respeita os filtros gerais da tela) e mostra `linhaTabelaHTML()` — as **mesmas 10 colunas** da tabela do fim (função compartilhada). Os filtros dentro do modal são os **mesmos 6 da tela principal** (Buscar, Segmento, Tributação, Status, Documentação, Gerente de Contas — `#m-*`), reusando `filtrarConjunto()`, e agem só sobre esse subconjunto (pedido explícito do usuário: "os filtros no modal precisam ser iguais os da tela principal"). Fecha no X, no clique fora da caixa e no Esc. `abrirModal(registros, titulo, contexto)` / `renderizarModalTabela()` / `fecharModal()`. Os cards de **navegação** (grid de Departamento/Segmento, cards de Unidade) não ligam esses handlers — `renderizarDocGrupo()` emite os `data-doc`/`data-status` sempre, mas só `renderizarQuebraGrupo()` liga os cliques por cima.

### Cards totalizadores clicáveis ("tabela dinâmica", 2026-08-25)
Todo quebra-card "Por Tributação" é clicável e mostra os clientes daquele valor **na tabela logo abaixo** — não só nos KPIs/cards, que era o comportamento antigo. Os cards de navegação (Unidade/Departamento) **não** usam esse mecanismo — o clique neles navega pra outra tela (ver seção de Navegação no topo), que já mostra a tabela filtrada por aquele escopo. O ranking por Gerente **não filtra mais a tabela** — desde 2026-09-02 abre o modal (acima).

**Clique no quebra-card é sempre um único alvo (2026-08-26, correção pedida pelo usuário)**: o card inteiro — incluindo o detalhamento de Documentação/Status dentro dele — só filtra pelo campo do próprio card (ex. Tributação = "Lucro Real"). Antes, o doc-grupo e cada linha de Status dentro do card tinham clique próprio que *combinava* com o campo do card (ex. clicar em "Com o GC" dentro do card "Lucro Real" filtrava Tributação=Lucro Real **+** Status=Com o GC) — o usuário achou essa combinação confusa (parecia que o clique tinha "ligado" um filtro que não era o pedido) e pediu que qualquer clique dentro do card ative só o filtro do card. `renderizarDocGrupo()` não tem mais parâmetro `clicavel`/handler próprio; o clique é só do `.quebra-card` externo (`alternarFiltroEMostrarTabela`), que já borbulha naturalmente de qualquer ponto interno do card. As classes CSS `.doc-grupo-clicavel`/`.status-linha-clicavel` e a função `ativarCliqueDetalhado()` foram removidas (ficaram sem uso).

Dois princípios do quebra-card clicável, o 2º ajustado depois de o usuário notar uma confusão real ao vivo (2026-08-25 — clicar em "Simples Nacional" parecia "ligar" o Status "Fechado" sozinho):
1. **Um clique sempre limpa todos os outros filtros gerais antes de aplicar o campo daquele card** (`limparFiltrosGerais()` dentro de `alternarFiltroEMostrarTabela()`) — nunca acrescenta em cima do que já estava ativo.
2. **Os filtros da tabela são sempre sincronizados por inteiro com os gerais** (`sincronizarFiltroTabelaComGeral()`, os 6 campos, não só o do clique) — nunca só copiados campo a campo. Sem isso, um filtro da tabela deixado de um clique anterior (num campo diferente) ficava "esquecido" e se combinava com o novo, podendo zerar a tabela mesmo com o card mostrando um número positivo.

`alternarFiltroEMostrarTabela(campo, valor)` alterna 1 campo (usado pelo quebra-card, ex. Tributação). `filtrarPorVariosEMostrarTabela` foi **removida** (2026-09-02) — só o ranking por Gerente a usava, e o ranking passou a abrir o modal em vez de filtrar a tabela.

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
