# CLAUDE.md — Regras de Projeto

## O que é este projeto

Portal único que junta o [[project_radar_fiscal]] e a [[project_analise_balanco]] numa página só, com uma aba grande no topo pra alternar entre "Radar Fiscal" e "Análise de Balanço" — mesma ideia da [[project_analise_entrega_sped]] (ICMS/Contribuições), aplicada aqui a duas fontes com robôs e esquemas de planilha totalmente diferentes (não só valores diferentes de um mesmo campo).

**Este projeto não tem robô de automação próprio** (não fala com MGApps nem Intranet diretamente) — mas tem um **orquestrador** (`backend/orquestrador.py`) que roda os dois robôs originais em sequência, carregando cada módulo pelo caminho (mesma técnica do hub [[project_atualizacao_de_bases]] em `backend/relatorios/*.py`):
- `Radar-Fiscal/backend/radar_fiscal.py` (MGApps desktop, pywinauto)
- `Analise-de-Balanco/backend/orquestrador.py` (MGApps + Intranet, pywinauto + selenium)

O portal continua sendo uma fusão no nível do **frontend** (cada fonte mantém seu próprio schema até a normalização em `static/script.js`), mas a **atualização dos dados agora é um só comando aqui** — `python backend/orquestrador.py` chama os dois robôs, copia os JSONs que cada um gera pro portal e ainda junta os dois resumos num **Excel único** (`data/relatorio_fechamentos/resumo.xlsx`, schema normalizado, pedido explícito do usuário — "quero os dados em excel mesmo"). Os dois projetos-fonte continuam existindo e podendo ser usados/rodados sozinhos — este orquestrador só os invoca, não substitui o código deles.

## Estrutura de Diretórios

```
Raiz/
├── index.html
├── .gitignore
├── CLAUDE.md
├── requirements.txt
├── static/
│   ├── script.js
│   ├── style.css   (copiado do Radar Fiscal, com as classes de aba renomeadas p/ "tipo-relatorio-*")
│   └── logo.png
├── data/
│   └── relatorio_fechamentos/
│       ├── radar_fiscal_dados.json       (cópia do Radar-Fiscal/data/radar_fiscal/, schema bruto — portal normaliza no frontend)
│       ├── status_radar_fiscal.json      (idem, renomeado)
│       ├── analise_balanco_dados.json    (cópia do Analise-de-Balanco/data/analise_balanco/)
│       ├── status_analise_balanco.json   (idem, renomeado)
│       ├── status.json                   (execução combinada — total das duas fontes)
│       └── resumo.xlsx                   (as duas fontes juntas, schema normalizado — gerado só pelo orquestrador, não versionado)
└── backend/
    └── orquestrador.py   (roda os 2 robôs-fonte, copia os JSONs, gera o Excel único)
```

---

## Regras — OBRIGATÓRIO SEGUIR

### HTML / CSS / JS
Mesmas regras dos outros portais MG: único `.html` é `index.html` na raiz; `.css`/`.js` só em `static/`; nada de lógica de backend em `static/`.

### Dados (`data/relatorio_fechamentos/`)
- **Nunca editar os JSONs nem o Excel à mão.** Todos são gerados por `backend/orquestrador.py` — a fonte da verdade real continua em `Radar-Fiscal/data/radar_fiscal/` e `Analise-de-Balanco/data/analise_balanco/`.
- Pra atualizar com dados novos: `python backend/orquestrador.py` aqui (assume que `Radar-Fiscal/` e `Analise-de-Balanco/` são pastas irmãs desta, mesmo nível — `RAIZ.parent`). Roda os dois robôs de verdade (MGApps + Intranet) — leva o tempo normal de cada um somado, e precisa rodar numa máquina com MGApps instalado e as credenciais configuradas nos `.env` dos dois projetos-fonte (este projeto não tem `.env` próprio).
- Os 4 JSONs (`radar_fiscal_dados.json`, `status_radar_fiscal.json`, `analise_balanco_dados.json`, `status_analise_balanco.json`) + `status.json` combinado são **versionados** no git (mesmo padrão dos outros portais — permite GitHub Pages mostrar dados atualizados a cada push). **`resumo.xlsx` não é versionado** (`.gitignore`) — é um artefato grande e binário, regenerado a cada execução do orquestrador; quem quiser os dados em Excel roda o orquestrador e abre o arquivo local.

### Normalização das duas fontes (`static/script.js`)
Radar Fiscal e Análise de Balanço têm nomes de coluna diferentes pra conceitos equivalentes. `normalizarRadarFiscal()`/`normalizarAnaliseBalanco()` traduzem cada fonte pro mesmo esquema comum **uma vez**, no carregamento (`carregarDados()`), pra todo o resto do arquivo (filtros, cards, ranking, evolução, tabela) trabalhar só com os nomes normalizados:

| Campo comum | Radar Fiscal | Análise de Balanço |
|---|---|---|
| `Id` | `IdCorporativo` | `IdCliente` |
| `Cliente` | `Nome` | `Cliente` |
| `Grupo` | `Grupo` | `Grupo` |
| `Unidade` | `Unidade` | `Unidade` |
| `Segmento` | `Segmento` | `Segmento` |
| `Gerente` | `GerenteContas` | `Gerente` |
| `Tributacao` | `RegimeApuracao` | `Tributacao` |
| `Status` | `Status` | `Status` (grafia própria) |
| `Documentacao` | `Documentação` | `Documentação` |
| `Departamento` | `DeptoFiscal` | *(não existe)* |
| `DataReferencia` | `DataConfirmacao` | `DataImportacao` |

`Segmento` e `Documentação` já têm o mesmo nome e os mesmos 2 valores nas duas fontes — não precisam de mapeamento, só de "carregar como está".

### Tipo de Relatório (aba do topo) — igual ao padrão do SPED
A página inteira (chips de Unidade, filtros, cards, ranking, evolução, tabela) sempre mostra só um Tipo de Relatório por vez (`tipoRelatorioAtivo`), escolhido em `#tipo-relatorio-abas`. Trocar de aba (`selecionarTipoRelatorio()`) reseta todos os filtros e repopula cada `<select>` só com os valores daquela fonte (`repopularSelect()`).

Duas diferenças reais entre as fontes que a UI precisa esconder/adaptar ao trocar de aba:
1. **Departamento só existe no Radar Fiscal.** Os grupos `#f-depto-grupo`/`#t-depto-grupo` recebem a classe `.oculto` quando o Tipo ativo é "Análise de Balanço" (não é só deixado vazio — some da tela). A coluna "Departamento" da tabela **fica sempre visível** (mostra "—" nas linhas da Análise de Balanço) — decisão consciente pra não arriscar desalinhar `<colgroup>`/`<td>` escondendo célula por célula (`display:none` num `<td>` isolado quebra a contagem de colunas do `table-layout: fixed`).
2. **A 2ª aba de cards** é "Por Departamento" (Radar Fiscal, aninhado por Tributação, usa `el.depto`) ou "Por Segmento" (Análise de Balanço, aninhado por Tributação, usa o mesmo `el.segmento` do filtro geral — igual ao próprio portal original da Análise de Balanço já fazia). Config em `QUEBRA_CONFIG_POR_TIPO`; o rótulo do botão (`#quebra-aba-segunda`) é atualizado por `selecionarTipoRelatorio()`.
3. **Os 5 valores reais de `Status`** são diferentes nas duas fontes, cada uma com uma categoria sem equivalente na outra (Radar Fiscal: "Bloqueado"; Análise de Balanço: "Importado Contábil") — por isso `STATUS_ORDEM_POR_TIPO` é por Tipo, não uma lista global única (evitaria mostrar uma linha fixa tipo "Importado Contábil: 0" em todo card do Radar Fiscal, que nunca deixaria de ser zero). `STATUS_ROTULOS_POR_TIPO` só ajusta a grafia da Análise de Balanço pra bater com o Radar Fiscal onde o significado é o mesmo (`"OK - Com GC"→"Com o GC"`, `"Não Importado"→"Não importado"`) — mesma regra que já existia no portal original da Análise de Balanço, só que agora vale também pra Status usado como filtro/rótulo de tabela (`rotuloStatus()`), não só nos cards.

O resto (ranking por Gerente, "Evolução Diária" por `DataReferencia`, tabela, cards "Por Tributação") é **código idêntico** pras duas fontes — só funciona porque os nomes já estão normalizados.

### `backend/orquestrador.py` — detalhes
- Carrega `Radar-Fiscal/backend/radar_fiscal.py` e `Analise-de-Balanco/backend/orquestrador.py` via `importlib.util.spec_from_file_location` (não são pacotes instaláveis, são projetos-irmãos fora deste repo) e chama o `executar(log=None) -> pd.DataFrame` de cada um, na ordem Radar Fiscal → Análise de Balanço. Cada robô continua gravando seus próprios arquivos brutos/`.xlsx`/JSON no projeto-fonte, normalmente — este orquestrador não muda o comportamento deles, só reaproveita o `DataFrame` que os dois já retornam.
- **Excel único** (`_normalizar_radar_fiscal()`/`_normalizar_analise_balanco()`): mesmo mapeamento de coluna que `normalizarRadarFiscal()`/`normalizarAnaliseBalanco()` em `static/script.js` (ver tabela acima), só que em pandas. Lê os nomes de coluna reais do `resumo.xlsx` de cada fonte — atenção que o Radar Fiscal usa `"Gerente de Contas"` (com espaços) no resumo, mesmo a coluna virando `GerenteContas` só no JSON do portal (rename feito dentro do próprio `radar_fiscal.py`, `_gerar_json_portal`); se o schema de qualquer um dos dois `resumo.xlsx` mudar, atualizar os nomes aqui.
- **JSONs do portal**: não são gerados a partir do Excel normalizado — são só **copiados** dos arquivos que cada robô-fonte já produz (`_copiar()`), no schema bruto de cada um. Por isso `static/script.js` não precisou mudar nada ao adicionar o orquestrador — continua normalizando os dois JSONs brutos como sempre fez.
- `status.json` (combinado, total das duas fontes) é gerado aqui; `status_radar_fiscal.json`/`status_analise_balanco.json` são cópias dos `status.json` de cada fonte (contagem/timestamp própria de cada robô) — o portal mostra o da aba ativa (`statusPorTipo`/`atualizarStatusExibido()` em `static/script.js`).
- Validado com `pandas` direto contra os `resumo.xlsx` já existentes dos dois projetos-fonte (sem rodar os robôs de verdade): 2074 Radar Fiscal + 2112 Análise de Balanço = 4186 linhas, colunas e nulos batendo com o esperado (`Departamento` nulo só nas linhas da Análise de Balanço + alguns registros do próprio Radar Fiscal sem `DeptoFiscal`). **A execução real dos dois robôs (MGApps + Intranet) não foi testada por mim** — precisa rodar numa máquina com MGApps instalado; robôs individualmente já validados nos projetos-fonte.

### Como rodar
```
cd Relatorio-de-Fechamentos
python backend/orquestrador.py   # roda os 2 robôs-fonte + gera Excel único + copia JSONs do portal
python -m http.server 8794
```
Acessar `http://localhost:8794`.

### Cores — regra fixa
**Nunca verde/âmbar para status.** Só rampa de vermelho MG, igual a todos os outros dashboards MG — ver [[feedback_mg_dashboards_red_only_palette]].

### Se um dia quiser automatizar via hub
Hoje `backend/orquestrador.py` é rodado manualmente. Se isso incomodar, o próximo passo natural é registrar este portal como mais um item do [[project_atualizacao_de_bases]] (hub que já roda os dois robôs-fonte separadamente) — um wrapper ali chamaria este orquestrador (ou a lógica equivalente) antes de dar commit+push. Não implementado ainda — não fazer sem o usuário pedir.
