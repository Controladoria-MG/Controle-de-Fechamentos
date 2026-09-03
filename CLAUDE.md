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
    ├── radar_fiscal.py            (robô 1 — MGApps "Sistema de Analise", pywinauto)
    ├── radar_fechamento.py        (robô 2a — MGApps "Analise Balanço", pywinauto)
    ├── retorno_checklist.py       (robô 2b — Intranet, selenium — flag Documentação Pendente/Recebida)
    ├── checklist_ctb_extrator.py  (robô 2c — Intranet, selenium — baixa o "Checklist Contábil > Recebimento")
    ├── checklist_em_aberto.py     (robô 2d — Intranet, selenium — "(Meu)checklist contabil em aberto")
    ├── checklist_ctb.py           (consolida o relatório CTB: 1 linha por cliente, motivos concatenados)
    ├── resumo.py                  (junta radar_fechamento + retorno_checklist)
    └── orquestrador.py            (roda tudo, gera o Excel único e os JSONs do portal)
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
CHECKLIST_CTB_USUARIO=  # conta do app "Checklist Contábil" (checklist_ctb_extrator.py); vazio = usa INTRANET_*
CHECKLIST_CTB_SENHA=
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
| `DocumentoPendente` | `"Planilha de Mercados.PENDENCIAS FECHAMENTO"` (máscara `ÍA=="A"`) | Checklist Contábil consolidado, junção por `IdCliente` |

### Coluna "Documento Pendente" — fontes diferentes por lado (2026-09-02)
É a última coluna da tabela / do modal. As duas fontes preenchem de lugares diferentes; mesmo destino (`DocumentoPendente` no schema normalizado).

**Radar Fiscal**: coluna `"PENDENCIAS FECHAMENTO"` da Planilha de Mercados (só SP/GOIAS passam por ela, `UNIDADES_MERCADOS`). Máscara `"ÍA" == "A"` em `radar_fiscal.py::_processar_resumo` — só as linhas com ÍA=="A" ficam com texto (pedido do usuário 2026-08-25, **re-confirmado 2026-09-02** — ele chegou a cogitar trocar pra `DOC. PENDENTE CHECKLIST` e voltou atrás; a planilha tem as duas colunas). Rename pra `DocumentoPendente` em `_gerar_json_portal`.

**Análise de Balanço**: combinação de DOIS relatórios da Intranet, além do `retorno_checklist.py` que já existia.
- `backend/checklist_ctb_extrator.py` — baixa "Checklist Contábil > Relatórios > Recebimento" (uma linha por documento pendente de cada cliente) pra `data/analise_balanco/checklist_ctb.xlsx`. Adaptado do `CTB.py` que o usuário largou na raiz (fixes: `--headless=new`, `.env`, competência como parâmetro, download pra `data/`). Conta: `CHECKLIST_CTB_USUARIO/SENHA`, cai pra `INTRANET_*`.
- `backend/checklist_em_aberto.py` — baixa "(Meu)checklist contabil em aberto" (MG Controle > Relatórios > Personalizados). Diz quais clientes têm tarefa de retorno do checklist EM ABERTO. Conta: `CHECKLIST_ABERTO_USUARIO/SENHA` — **precisa ser a conta pessoal `warruda`** (relatório "(Meu)...", só aparece na lista de Personalizados do próprio dono). Expõe `clientes_em_aberto() -> set[int]`.
- `backend/checklist_ctb.py` — consolida o relatório de Recebimento em 1 linha por `IdCliente`, motivos concatenados no **formato "Tipo agrupado"** (escolha do usuário): um `Tipo` por linha; várias `Descricao` do mesmo Tipo viram lista com `; `; descrição vazia/igual ao Tipo é descartada.
- `orquestrador.py::_juntar_checklist_ctb` — LEFT JOIN por `IdCliente`, mas **só preenche `DocumentoPendente` pros clientes que estão em `clientes_em_aberto()`** (regra do usuário 2026-09-02: "só quando a tarefa de retorno do checklist estiver em aberto"). Sem o relatório de em-aberto, a coluna fica vazia (não dá pra aplicar a regra). Nenhum dos dois arquivos ausente derruba o pipeline (try/except no `executar()`, dentro de um loop pelos 2 extratores).

**Estado da validação ao vivo (2026-09-03):**
- `checklist_ctb_extrator.py` — **✅ rodou de primeira** (2026-09-02) com o fallback `INTRANET_*`. Baixou 8.953 linhas (aba "Pendências CTB", 19 colunas), `checklist_ctb.py` consolidou pra 1.502 clientes. Navegação toda validada.
- `checklist_em_aberto.py` — **✅ VALIDADO (2026-09-03)** com `CHECKLIST_ABERTO_USUARIO=warruda` no `.env`. Competência 08/2026 → 754 tarefas "Retorno do Check-List" em aberto, 754 `CodCliente` únicos. Os 3 chutes bateram: `XPATH_MENU_RELATORIO` casou o item, aba = `"Pendencias"` (o xlsx tem também `"Totalizador"`), coluna-chave = `"CodCliente"`. `clientes_em_aberto()` retorna as 754.
- `orquestrador.py` completo — **✅ RODOU PONTA A PONTA (2026-09-03, 3ª tentativa)**. As 2 primeiras falharam no login do "Sistema de Analise" (`SENHA` do `.env` estava desatualizada — o usuário tinha editado `MGAPPS_SENHA`, que **nenhum código lê**; o certo é a linha `SENHA=`). Resultado: 2.095 Radar Fiscal + 2.134 Análise de Balanço = 4.229; `resumo.xlsx` gerado; JSONs do portal regenerados com `DocumentoPendente`. Checklist CTB: 1.495 no recebimento, 754 em aberto, **740 preenchidos** na AB.

**Exibição LIGADA no frontend (2026-09-03)**: `normalizarAnaliseBalanco` em `static/script.js` usa `DocumentoPendente: formatarDocumentoPendente(r.DocumentoPendente)`.

**Ainda NÃO commitado** (2026-09-03) — depois da reforma de "Documentação Recebida sem tarefa em aberto" (ver seção acima), regenerei só o `analise_balanco_dados.json` a partir dos `.xlsx` já baixados pela execução completa (sem reabrir automação). Falta: validar visualmente, rodar o `orquestrador.py` completo mais uma vez (pra o `resumo.xlsx` também pegar a regra nova) OU aceitar que o Excel só normaliza na próxima execução, e commitar tudo junto. `.gitignore` já cobre `/CTB.py`, `/Baixar_Checklist_CTB.py`, `/relat*.xlsx`, `/Checklist_Contabil_*.xlsx`, `_temp_ctb/`, `_temp_aberto/`.

`formatarDocumentoPendente()` no frontend só remove espaço nas pontas.

`Segmento` e `Documentação` já têm o mesmo nome e os mesmos 2 valores nas duas fontes — não precisam de mapeamento. **Atenção**: o Radar Fiscal usa `"Gerente de Contas"` (com espaço) no `resumo.xlsx`, e só vira `GerenteContas` no JSON do portal (rename feito dentro do próprio `radar_fiscal.py`, `_gerar_json_portal`) — `orquestrador.py` lê direto do `resumo.xlsx`/DataFrame retornado por `radar_fiscal.executar()`, então usa o nome com espaço.

### Tipo de Relatório (aba do topo) — igual ao padrão do SPED
A página inteira (navegação de Unidade/Departamento, filtros, cards, ranking, evolução, tabela) sempre mostra só um Tipo de Relatório por vez (`tipoRelatorioAtivo`), escolhido em `#tipo-relatorio-abas`. Trocar de aba (`selecionarTipoRelatorio()`) volta pro Painel de Unidades, reseta todos os filtros e repopula cada `<select>` só com os valores daquele escopo (`repopularSelect()`).

Diferenças reais entre as fontes que a UI precisa esconder/adaptar ao trocar de aba:
1. **Departamento só existe no Radar Fiscal.** É a "2ª dimensão" de navegação (`QUEBRA_CONFIG_POR_TIPO[tipo].segunda`) — "Por Departamento" no Radar Fiscal, "Por Segmento" na Análise de Balanço (reaproveita a mesma coluna do filtro geral "Segmento"). O campo `Departamento` continua no schema normalizado (`normalizarRadarFiscal`, de `DeptoFiscal`) só pra essa navegação. **A coluna "Departamento" foi removida da tabela e do modal (2026-09-03)** — o usuário pediu ("tire essa coluna") porque ficava sempre "—" nas linhas da Análise de Balanço. Tabela/modal agora têm **9 colunas** (Cliente, Grupo, Unidade, Segmento, Gerente de Contas, Tributação, Status, Documentação, Documento Pendente); `<colgroup>` dos dois ajustado pra somar 100%.
2. **Os 5 valores reais de `Status`** são diferentes nas duas fontes, cada uma com uma categoria sem equivalente na outra (Radar Fiscal: "Bloqueado"; Análise de Balanço: "Importado Contábil") — por isso `STATUS_ORDEM_POR_TIPO` é por Tipo, não uma lista global única. `STATUS_ROTULOS_POR_TIPO` só ajusta a grafia da Análise de Balanço pra bater com o Radar Fiscal onde o significado é o mesmo (`"OK - Com GC"→"Com o GC"`, `"Não Importado"→"Não importado"`), também usada no filtro/rótulo de tabela (`rotuloStatus()`), não só nos cards. O último item de cada lista em `STATUS_ORDEM_POR_TIPO` é sempre o equivalente a "não importado" — usado por `renderizarDocGrupo()` (ver regra de Documentação abaixo).

O resto (ranking por Gerente, "Evolução Diária" por `DataReferencia`, tabela, cards "Por Tributação") é **código idêntico** pras duas fontes — só funciona porque os nomes já estão normalizados.

### Documentação Pendente só deveria coexistir com Status "Não importado" (2026-08-25)
Pedido do usuário: "todas as empresas [com documentação pendente] precisam estar com um status Não importado" — na prática, checando os dados reais, isso não era verdade (Status "Fechado" era o caso mais comum, mas "Simulando"/"Com o GC" também apareciam com documentação pendente — mais de 700 das 4190 linhas na calibração de 2026-08-24). Corrigido no **tratamento da base**, não só na exibição — qualquer linha com `Documentação == "Documentação Pendente"` e `Status` diferente do "não importado" daquela fonte é normalizada para `"Documentação Recebida"`:
- Frontend: `corrigirDocumentacao()` em `static/script.js`, chamada de dentro de `normalizarRadarFiscal()`/`normalizarAnaliseBalanco()` (afeta a exibição do portal).
- Backend: `_corrigir_documentacao_inconsistente()` em `backend/orquestrador.py`, chamada depois do `pd.concat()` dos dois normalizados (afeta o `resumo.xlsx`).
- Os JSONs brutos (`radar_fiscal_dados.json`/`analise_balanco_dados.json`) **não** são alterados — continuam no schema bruto de cada fonte, como o resto da normalização (o frontend é quem corrige na hora de montar `dados`).

Como consequência, `renderizarDocGrupo()` simplifica o card de "Documentação Pendente" pra mostrar só 1 linha de status (a "não importado" da fonte) em vez das 5 — as outras seriam sempre zero.

### Análise de Balanço: a Documentação é decidida pela tarefa de retorno do checklist em aberto (2026-09-03)
Pedidos do usuário, dois no mesmo dia:
1. "não tem como uma tarefa pendente não ter documentação — quando não houver tarefa de retorno do checklist, considera a documentação recebida".
2. (ao ver o cliente 6989 ATRIA, Status "Fechado", com tarefa em aberto, aparecendo como "Recebida") "como a tarefa está como recebida e tem documentação pendente? essa tarefa está em aberta na base".

Antes a AB marcava `"Documentação Pendente"` só por o Status ser `"Não Importado"`. Agora a **`Documentação` da Análise de Balanço = a tarefa de retorno do checklist**:
- `IdCliente` ∈ `checklist_em_aberto.clientes_em_aberto()` → `"Documentação Pendente"` (independente do Status — a tarefa está aberta).
- `IdCliente` ∉ em aberto → `"Documentação Recebida"`.

Aplicado em `_juntar_checklist_ctb()` (junto com o `DocumentoPendente`), então **entra no `analise_balanco_dados.json` do portal** e flui pro `resumo.xlsx`. Efeito 2026-09-03: 944 linhas viraram Recebida; `Documentação Pendente` = exatamente os ~747 com tarefa em aberto. Só vale quando o relatório "em aberto" existe (mesma guarda do `DocumentoPendente`).

**A regra de Status (2026-08-25) não vale mais pra AB.** `corrigirDocumentacao()` no frontend agora **só é chamada em `normalizarRadarFiscal`** — `normalizarAnaliseBalanco` usa `r["Documentação"]` direto (o backend já resolveu). `_corrigir_documentacao_inconsistente()` no backend ganhou um filtro `TipoRelatorio == "Radar Fiscal"`. Sem isso, a regra de Status re-quebrava os ~5 clientes em aberto com Status ≠ "Não Importado" (2 Com GC / 2 Fechado / 1 Excluído), voltando a mostrá-los como "Recebida" — exatamente a queixa do cliente 6989.

`renderizarDocGrupo()` no card "Documentação Pendente" agora mostra a linha "não importado" **mais** qualquer outro Status com contagem > 0 (antes era só a "não importado", que pro Radar Fiscal segue sendo a única não-zero).

- **Casos que ainda mostram "Pendente" + "—"** (Documento Pendente vazio): ~7 clientes que TÊM tarefa em aberto mas não têm nenhuma linha no relatório de Recebimento. A regra do usuário não cobre esses (eles têm tarefa). Deixados como estão.

### `backend/orquestrador.py` — fluxo
1. `_recarregar_credenciais()` — relê `.env`, sobrescreve `radar_fiscal.USUARIO`/`SENHA` e `retorno_checklist.USUARIO`/`SENHA`.
2. Checagem de janelas abertas (MG Apps + Análise de Balanço) de uma execução anterior.
3. `radar_fiscal.executar(log)` → `df_radar`. Fecha MG Apps de novo.
4. `radar_fechamento.executar(log)` → `arquivo_radar_fechamento`; `retorno_checklist.executar(log)` → `arquivo_checklist`; `resumo.gerar_resumo(...)` → `df_balanco`. Depois `checklist_ctb_extrator.executar(log)` baixa o Checklist Contábil (**try/except — não bloqueia o pipeline** se cair) e `_juntar_checklist_ctb()` faz o LEFT JOIN da coluna `DocumentoPendente` por `IdCliente`. Gera `analise_balanco_dados.json`/`status.json` em `data/analise_balanco/` (réplica do que o antigo `Analise-de-Balanco/backend/orquestrador.py` fazia — este projeto não tem mais aquele orquestrador intermediário).
5. Copia os JSONs/status de `data/radar_fiscal/` e `data/analise_balanco/` pra `data/relatorio_fechamentos/`.
6. Junta os dois (normalizados) num `resumo.xlsx` único.
7. Escreve `status.json` combinado.

### Placares e abas Pendente/Concluído seguem a coluna Documentação (2026-09-03)
**Histórico**: em 2026-09-02 os placares passaram a medir o FECHAMENTO (`Status === "Fechado"`) em vez da documentação, a pedido do usuário. Em 2026-09-03, com a coluna "Documentação" da Análise de Balanço virando um sinal confiável (tarefa de retorno do checklist em aberto), o usuário pediu que **a aba da tabela sempre bata com a coluna Documentação da linha** ("quando for pendente, esteja na aba pendente; quando for concluído/recebido, na aba concluída — não tem segredo"). Então voltaram a seguir a documentação:

Os 3 placares do topo (`renderizarPlacares`): **Total de Empresas / Documentação Recebida / Documentação Pendente**. O corte é `documentacaoRecebida(r)` = `r.Documentacao === "Documentação Recebida"` (schema já normalizado). `fechamentoConcluido()`/`STATUS_CONCLUIDO` **foram removidos**. Recebida/Pendente são clicáveis: trocam a aba da tabela do fim (mesmo `data-aba` "Concluído"/"Pendente" — só os rótulos dos placares mudaram) e rolam até ela.

`#tabela-abas`: `abaTabelaAtiva` ("Pendente"|"Concluído") é um corte em `aplicarFiltroTabela()` **depois** dos 6 filtros: Pendente = `!documentacaoRecebida(r)`, Concluído = `documentacaoRecebida(r)`. Como o placar e a aba usam o mesmo predicado, clicar num placar sempre cai numa aba com o mesmo número. `definirAbaTabela()` (reset, volta pra "Pendente") vs `trocarAbaTabela()` (troca + re-renderiza). Os `data-aba` internos continuam "Pendente"/"Concluído"; os `<button>` no HTML seguem rotulados "Pendente"/"Concluído".

### Modal de registros (2026-09-02) — IDÊNTICO ao portal de Análise de Entrega de SPED
Pedido literal do usuário: "o modal precisa ser exatamente igual o do Análise e Entrega de SPED". HTML/CSS/JS do modal copiados 1:1 de `Analise-de-Entrega-de-SPED/` — mesma `.modal-caixa`/`.modal-cabecalho`/`.modal-filtros` (flexbox plano, sem `.filtros-grid`/`.filtro-label` — só `<input>` + `<select>`s diretos, com placeholders descritivos "Todos os segmentos" etc.), mesmo `.modal-tabela th` (fundo cinza-claro `--bg-card-alt`, uppercase, NÃO vermelho), mesmo `.status-linha.linha-modal:hover` (texto fica vermelho, sem tint de fundo — igual `.status-linha-estagio` lá). Ao replicar o modal do SPED aqui, adaptar só o que é específico do domínio: os **6 filtros** (Buscar, Segmento, Tributação, Status, Documentação, Gerente — `#m-*`) e as **9 colunas** (as da tabela do fim deste projeto, via `linhaTabelaHTML()` compartilhada — "Departamento" foi removida em 2026-09-03) — cada modal espelha a tabela principal do seu próprio portal.

`#modal-registros` abre ao clicar numa **linha de Status** (`.status-linha`, com `stopPropagation`) dentro de:
- um card **"Por Tributação"** (`renderizarQuebraGrupo`);
- um card **"Por Departamento" / "Por Segmento"** (grid de navegação, `renderizarCardsNavegacao`) — adicionado 2026-09-03 a pedido do usuário ("coloque o modal no restante", não só nos "Por Tributação"). O clique no corpo do card continua navegando; só a `.status-linha` abre o modal;
- uma **linha do ranking por Gerente**.

A ligação dos cliques de `.status-linha` é a função compartilhada **`ligarModalNosCards(container, rows, chave)`** (= `ligarCliquesMotivo()` do SPED) — chamada por `renderizarQuebraGrupo` (com `filtrados`/`campo`) e por `renderizarCardsNavegacao` (com as linhas da unidade / `chave` da 2ª dimensão). O cabeçalho `.doc-cabecalho` **não** é clicável (idem SPED). Cards de **Unidade** (tela 1, `renderizarCardsUnidades`) não têm detalhamento de Status — nada a ligar, igual ao SPED. `abrirModal(registros, titulo, contexto)` / `renderizarModalTabela()` / `fecharModal()`; fecha no X, clique fora e Esc.

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
