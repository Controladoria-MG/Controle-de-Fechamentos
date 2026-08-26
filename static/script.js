// Portal único que junta o Radar Fiscal e a Análise de Balanço (mesma ideia
// da Análise de Entrega de SPED com ICMS/Contribuições): os dois relatórios
// têm robôs e planilhas totalmente diferentes, então cada fonte é
// normalizada pra um esquema comum (ver normalizarRadarFiscal/
// normalizarAnaliseBalanco) antes de entrar em `dados`. A partir daí, a
// página inteira (filtros, cards, ranking, evolução, tabela) sempre mostra
// só um Tipo de Relatório por vez, escolhido na aba do topo — igual ao
// padrão do SPED.
//
// Navegação (Unidade -> Departamento/Segmento) segue o mesmo padrão de
// telas do Portal de Tarefas: escolher o card da Unidade abre a tela da
// Unidade (visão Consolidada da unidade inteira + cards de Departamento/
// Segmento pra detalhar); escolher um desses cards abre a tela do
// Departamento (mesmo corpo do dashboard, mais restrito). `escopo` guarda
// esse estado de navegação; `dadosEscopo` é `dadosTipo` já recortado por ele.
let dados = [];
let dadosTipo = [];
let dadosEscopo = [];
let tipoRelatorioAtivo = null;
let escopo = { unidade: null, depto: null };
let filtrados = [];
let filtradosTabela = [];

const el = {
  status: document.getElementById("status-execucao"),
  tipoRelatorioAbas: document.getElementById("tipo-relatorio-abas"),
  breadcrumbBar: document.getElementById("navegacao-breadcrumb"),
  breadcrumbCrumbs: document.getElementById("breadcrumb-crumbs"),
  btnVoltarPainel: document.getElementById("btn-voltar-painel"),
  secaoUnidades: document.getElementById("secao-unidades"),
  unidadesGrid: document.getElementById("unidades-grid"),
  secaoDepartamentos: document.getElementById("secao-departamentos"),
  departamentosGridTitulo: document.getElementById("departamentos-grid-titulo"),
  departamentosGrid: document.getElementById("departamentos-grid"),
  corpoDashboard: document.getElementById("corpo-dashboard"),
  tabelaSecao: document.getElementById("tabela-secao"),
  busca: document.getElementById("f-busca"),
  segmento: document.getElementById("f-segmento"),
  regime: document.getElementById("f-regime"),
  status_: document.getElementById("f-status"),
  documentacao: document.getElementById("f-documentacao"),
  gerente: document.getElementById("f-gerente"),
  limpar: document.getElementById("f-limpar"),
  corpo: document.getElementById("tabela-corpo"),
  contagem: document.getElementById("contagem"),
  quebraConteudo: document.getElementById("quebra-conteudo"),
  rankingGerentes: document.getElementById("ranking-gerentes"),
  evolucaoGrafico: document.getElementById("evolucao-grafico"),
  // Filtro independente, só da tabela — não afeta cards/ranking/evolução.
  tBusca: document.getElementById("t-busca"),
  tSegmento: document.getElementById("t-segmento"),
  tRegime: document.getElementById("t-regime"),
  tStatus: document.getElementById("t-status"),
  tDocumentacao: document.getElementById("t-documentacao"),
  tGerente: document.getElementById("t-gerente"),
  tLimpar: document.getElementById("t-limpar"),
};

// ── Normalização por fonte ──────────────────────────────────────────────
// Radar Fiscal e Análise de Balanço são robôs/planilhas independentes com
// nomes de coluna diferentes pra conceitos equivalentes (Nome/Cliente,
// GerenteContas/Gerente, RegimeApuracao/Tributacao, DataConfirmacao/
// DataImportacao...). Normalizar aqui, uma vez, no carregamento, permite
// que todo o resto do arquivo (filtros, cards, ranking, evolução, tabela)
// trabalhe só com os nomes comuns, sem `if (tipo === ...)` espalhado pelo
// código. `Segmento` e `Documentação` já têm o mesmo nome/valores nas duas
// fontes — não precisam de mapeamento.

// Pedido do usuário (2026-08-25): "todas as empresas [com documentação
// pendente] precisam estar com status Não importado". Ele só citou o caso
// "Fechado" como exemplo, mas nos dados reais qualquer Status diferente do
// "não importado" da fonte às vezes aparece com documentação pendente
// (Fechado é o mais comum, mas Simulando/Com o GC também acontecem) — pra
// cumprir a regra geral, tratamos qualquer uma dessas combinações como
// inconsistência de dados e normalizamos para "Documentação Recebida".
// Mesma regra replicada em
// backend/orquestrador.py::_corrigir_documentacao_inconsistente pro Excel.
function corrigirDocumentacao(status, documentacao, statusNaoImportado) {
  if (documentacao === "Documentação Pendente" && status !== statusNaoImportado) return "Documentação Recebida";
  return documentacao;
}

// A "Planilha de Mercados" traz um comentário livre sobre a pendência de
// fechamento de cada empresa (coluna "PENDENCIAS FECHAMENTO") — só existe
// pro Radar Fiscal (só SP/GOIAS passam pela Planilha de Mercados) e o
// backend já garante que só vem preenchido quando "ÍA" == "A" (pedido do
// usuário, 2026-08-25 — ver máscara em radar_fiscal.py::_processar_resumo).
function formatarDocumentoPendente(texto) {
  const limpo = (texto || "").trim();
  return limpo || null;
}

function normalizarRadarFiscal(r) {
  return {
    Id: r.IdCorporativo,
    Cliente: r.Nome,
    Grupo: r.Grupo,
    Unidade: r.Unidade,
    Segmento: r.Segmento,
    Gerente: r.GerenteContas,
    Tributacao: r.RegimeApuracao,
    Status: r.Status,
    Documentacao: corrigirDocumentacao(r.Status, r["Documentação"], "Não importado"),
    Departamento: r.DeptoFiscal,
    DataReferencia: r.DataConfirmacao,
    DocumentoPendente: formatarDocumentoPendente(r.DocumentoPendente),
    TipoRelatorio: "Radar Fiscal",
  };
}

function normalizarAnaliseBalanco(r) {
  return {
    Id: r.IdCliente,
    Cliente: r.Cliente,
    Grupo: r.Grupo,
    Unidade: r.Unidade,
    Segmento: r.Segmento,
    Gerente: r.Gerente,
    Tributacao: r.Tributacao,
    Status: r.Status,
    Documentacao: corrigirDocumentacao(r.Status, r["Documentação"], "Não Importado"),
    Departamento: undefined,
    DataReferencia: r.DataImportacao,
    DocumentoPendente: undefined,
    TipoRelatorio: "Análise de Balanço",
  };
}

function popularSelect(select, valores, formatar = (v) => v) {
  const atuais = new Set(Array.from(select.options).map((o) => o.value));
  [...valores].sort((a, b) => a.localeCompare(b, "pt-BR")).forEach((valor) => {
    if (!atuais.has(valor)) {
      const opt = document.createElement("option");
      opt.value = valor;
      opt.textContent = formatar(valor);
      select.appendChild(opt);
    }
  });
}

// Igual a popularSelect, mas limpa as opções antigas primeiro (mantendo só
// o placeholder "Todos"/"Todas", sempre a primeira <option>) — usado ao
// trocar de Tipo de Relatório/Unidade/Departamento, já que os valores
// possíveis de cada filtro mudam de um escopo pro outro.
function repopularSelect(select, valores, formatar = (v) => v) {
  const placeholder = select.options[0];
  select.innerHTML = "";
  select.appendChild(placeholder);
  popularSelect(select, valores, formatar);
}

function filtrarConjunto(conjunto, campos) {
  const busca = campos.busca.value.trim().toLowerCase();
  const segmento = campos.segmento.value;
  const regime = campos.regime.value;
  const status = campos.status.value;
  const documentacao = campos.documentacao.value;
  const gerente = campos.gerente.value;

  return conjunto.filter((r) => {
    if (segmento && r.Segmento !== segmento) return false;
    if (regime && r.Tributacao !== regime) return false;
    if (gerente && r.Gerente !== gerente) return false;
    if (status && r.Status !== status) return false;
    if (documentacao && r.Documentacao !== documentacao) return false;
    if (busca) {
      const alvo = `${r.Id ?? ""} ${r.Cliente || ""} ${r.Grupo || ""}`.toLowerCase();
      if (!alvo.includes(busca)) return false;
    }
    return true;
  });
}

function aplicarFiltros() {
  filtrados = filtrarConjunto(dadosEscopo, {
    busca: el.busca, segmento: el.segmento, regime: el.regime,
    status: el.status_, documentacao: el.documentacao, gerente: el.gerente,
  });

  renderizarQuebras();
  renderizarRankingGerentes();
  renderizarEvolucao();
}

function aplicarFiltroTabela() {
  filtradosTabela = filtrarConjunto(dadosEscopo, {
    busca: el.tBusca, segmento: el.tSegmento, regime: el.tRegime,
    status: el.tStatus, documentacao: el.tDocumentacao, gerente: el.tGerente,
  });
  renderizarTabela();
}

const ORDEM_DOCUMENTACAO = ["Documentação Recebida", "Documentação Pendente"];
// Os 5 valores reais de Status são diferentes entre as duas fontes — cada
// uma tem uma categoria sem equivalente na outra (Radar Fiscal: "Bloqueado";
// Análise de Balanço: "Importado Contábil"). Por isso a ordem/lista é por
// Tipo de Relatório, não global — um card do Radar Fiscal nunca mostra uma
// linha fixa de "Importado Contábil: 0" que não faz sentido pra ele. O
// último item de cada lista é sempre o equivalente a "não importado" —
// usado por renderizarDocGrupo pra simplificar o card de Documentação
// Pendente (ver ali).
const STATUS_ORDEM_POR_TIPO = {
  "Radar Fiscal": ["Fechado", "Bloqueado", "Simulando", "Com o GC", "Não importado"],
  "Análise de Balanço": ["Fechado", "Importado Contábil", "Simulando", "OK - Com GC", "Não Importado"],
};
// Só ajusta a grafia exibida da Análise de Balanço pra bater com o Radar
// Fiscal onde o significado é o mesmo (contagem interna continua pela
// chave real) — não inventa correspondência pra status sem equivalente.
const STATUS_ROTULOS_POR_TIPO = {
  "Radar Fiscal": {},
  "Análise de Balanço": { "OK - Com GC": "Com o GC", "Não Importado": "Não importado" },
};

function statusOrdem() {
  return STATUS_ORDEM_POR_TIPO[tipoRelatorioAtivo] || [];
}

function rotuloStatus(status) {
  const mapa = STATUS_ROTULOS_POR_TIPO[tipoRelatorioAtivo] || {};
  return mapa[status] || status;
}

function criarContadorStatus() {
  const status = new Map();
  statusOrdem().forEach((s) => status.set(s, 0));
  return { total: 0, status };
}

function criarContadorDocs() {
  const docs = new Map();
  ORDEM_DOCUMENTACAO.forEach((doc) => docs.set(doc, criarContadorStatus()));
  return docs;
}

function contarDetalhado(rows, chave) {
  const grupos = new Map();
  rows.forEach((r) => {
    const valor = r[chave];
    if (!valor) return;
    if (!grupos.has(valor)) grupos.set(valor, { total: 0, docs: criarContadorDocs() });
    const g = grupos.get(valor);
    g.total++;

    const doc = r.Documentacao || "Sem documentação";
    if (!g.docs.has(doc)) g.docs.set(doc, criarContadorStatus());
    const d = g.docs.get(doc);
    d.total++;

    const status = r.Status || "Não importado";
    d.status.set(status, (d.status.get(status) || 0) + 1);
  });
  return [...grupos.entries()].sort((a, b) => b[1].total - a[1].total);
}

function formatarPct(n) {
  return `${n.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%`;
}

// ── Cards totalizadores clicáveis ("tabela dinâmica") ───────────────────
// Cada dimensão mostrada nos cards (Tributação, Documentação, Status,
// Gerente) pode ser clicada pra filtrar a tabela pelos clientes daquele
// valor — sincroniza o filtro geral (que já alimentava só os KPIs/cards)
// com o filtro correspondente da tabela e rola a tela até ela.
const CAMPO_PARA_FILTROS = {
  Segmento: () => [el.segmento, el.tSegmento],
  Tributacao: () => [el.regime, el.tRegime],
  Status: () => [el.status_, el.tStatus],
  Documentacao: () => [el.documentacao, el.tDocumentacao],
  Gerente: () => [el.gerente, el.tGerente],
};

// Clicar em qualquer card/linha SEMPRE limpa os outros filtros gerais antes
// de aplicar o(s) campo(s) daquele clique — pedido explícito do usuário
// (2026-08-25), depois de notar que clicar em "Simples Nacional" parecia
// "ligar" o Status "Fechado" sozinho: o motivo era um filtro de Status
// deixado por um clique anterior (em outra parte da tela) que continuava
// ativo, já que a versão aditiva anterior só acrescentava o campo do clique
// sem mexer nos outros. Comportamento único e previsível agora: um clique
// sempre mostra exatamente (e só) o que aquele card/linha representa.
// O filtro da TABELA (independente por design, não afeta KPIs/cards) é
// sempre sincronizado por inteiro com o geral logo em seguida — nunca só
// o(s) campo(s) do clique — pra nenhum filtro da tabela "esquecido" de um
// clique anterior ficar combinado por engano com o novo.
function sincronizarFiltroTabelaComGeral() {
  el.tBusca.value = el.busca.value;
  el.tSegmento.value = el.segmento.value;
  el.tRegime.value = el.regime.value;
  el.tStatus.value = el.status_.value;
  el.tDocumentacao.value = el.documentacao.value;
  el.tGerente.value = el.gerente.value;
}

function alternarFiltroEMostrarTabela(campo, valor) {
  const [geralEl] = CAMPO_PARA_FILTROS[campo]();
  const estavaAtivo = geralEl.value === valor;
  limparFiltrosGerais();
  if (!estavaAtivo) {
    CAMPO_PARA_FILTROS[campo]()[0].value = valor;
  }
  sincronizarFiltroTabelaComGeral();
  aplicarFiltros();
  aplicarFiltroTabela();
  el.tabelaSecao.scrollIntoView({ behavior: "smooth", block: "start" });
}

// Igual a alternarFiltroEMostrarTabela, mas fixa (ou desliga, se já estava
// ativa a combinação inteira — mesmo toggle) vários campos de uma vez —
// usado pelos níveis "de dentro" de um card (doc-grupo/status-linha), que
// precisam combinar o próprio filtro com o do card "pai" (Tributação). Sem
// isso, clicar em "Documentação Pendente: 2" dentro do card "Lucro Real"
// filtrava só por Documentação, mostrando todas as pendentes de SP (7), não
// as 2 respectivas daquele card.
function filtrarPorVariosEMostrarTabela(pares) {
  const jaEstavaAtivo = pares.every(([campo, valor]) => CAMPO_PARA_FILTROS[campo]()[0].value === valor);
  limparFiltrosGerais();
  if (!jaEstavaAtivo) {
    pares.forEach(([campo, valor]) => {
      CAMPO_PARA_FILTROS[campo]()[0].value = valor;
    });
  }
  sincronizarFiltroTabelaComGeral();
  aplicarFiltros();
  aplicarFiltroTabela();
  el.tabelaSecao.scrollIntoView({ behavior: "smooth", block: "start" });
}

// Detalhamento de Documentação dentro de um card (quebra-card ou card de
// navegação) — não tem handler de clique próprio: o clique é sempre do
// card inteiro (ver quem chama esta função).
function renderizarDocGrupo(docNome, d, totalCategoria) {
  const classe = docNome === "Documentação Recebida" ? "recebida" : "pendente";
  const pctDoc = totalCategoria ? (d.total / totalCategoria) * 100 : 0;
  const chaveNaoImportado = statusOrdem()[statusOrdem().length - 1];
  // Documentação Pendente: depois da correção em corrigirDocumentacao()
  // (Fechado + Pendente vira Recebida), a única situação que sobra de
  // verdade é "Não importado" — listar os outros 4 status aqui seria
  // sempre zero e irrelevante, por isso mostra só essa linha.
  const statusOrdenado = classe === "pendente"
    ? [[chaveNaoImportado, d.status.get(chaveNaoImportado) || 0]]
    : statusOrdem().filter((s) => s !== chaveNaoImportado).map((s) => [s, d.status.get(s) || 0]);

  const linhasStatus = statusOrdenado
    .map(([status, count]) => {
      const pctStatus = totalCategoria ? (count / totalCategoria) * 100 : 0;
      const rotulo = rotuloStatus(status);
      return `
        <div class="status-linha">
          <span class="status-nome" title="${rotulo}">${rotulo}</span>
          <span class="status-valores"><b>${count.toLocaleString("pt-BR")}</b><span class="status-pct">${formatarPct(pctStatus)}</span></span>
        </div>
      `;
    })
    .join("");

  return `
    <div class="doc-grupo ${classe}">
      <div class="doc-cabecalho">
        <span class="doc-rotulo"><i class="ponto ${classe}"></i>${docNome}</span>
        <span class="doc-valores"><b>${d.total.toLocaleString("pt-BR")}</b><span class="doc-pct">${formatarPct(pctDoc)}</span></span>
      </div>
      <div class="status-lista">${linhasStatus}</div>
    </div>
  `;
}

// ── Cards de navegação (telas de Unidade e de Departamento/Segmento) ────
// Reaproveita a mesma aparência dos quebra-card (com o detalhamento de
// Documentação já pronto em renderizarDocGrupo), só que o clique no card
// inteiro navega pra outra tela em vez de filtrar.
function renderizarCardsNavegacao(container, rows, chave, aoClicar, mensagemVazio, formatarNome = (v) => v) {
  const grupos = contarDetalhado(rows, chave);
  if (!grupos.length) {
    container.innerHTML = `<p class="evolucao-vazio">${mensagemVazio}</p>`;
    return;
  }
  container.innerHTML = grupos
    .map(([nome, g]) => {
      const docsHtml = ORDEM_DOCUMENTACAO
        .map((docNome) => renderizarDocGrupo(docNome, g.docs.get(docNome), g.total))
        .join("");
      return `
        <div class="quebra-card nav-card" data-valor="${nome.replace(/"/g, "&quot;")}" title="${nome}">
          <div class="quebra-cabecalho">
            <div class="quebra-nome">${formatarNome(nome)}</div>
            <div class="quebra-total-num">${g.total.toLocaleString("pt-BR")}</div>
          </div>
          <div class="quebra-docs">${docsHtml}</div>
          <div class="nav-card-footer">Clique para ver os detalhes</div>
        </div>
      `;
    })
    .join("");

  container.querySelectorAll(".nav-card").forEach((cardEl) => {
    cardEl.addEventListener("click", () => aoClicar(cardEl.dataset.valor));
  });
}

// Tela 1 (Painel de Controle) — igual ao Portal de Tarefas: só o nome da
// unidade e um rodapé fixo, sem nenhum número/estatística no card.
// Ordem fixa pedida pelo usuário (2026-08-25) — não alfabética. Unidade que
// aparecer nos dados sem estar nesta lista vai pro final, em ordem alfabética.
// Grafia tem que bater exata com o dado bruto — "Santos" vem com essa
// capitalização mesmo (as outras 3 vêm em caixa alta), confirmado no JSON.
const ORDEM_UNIDADES = ["SP", "RJ", "Santos", "GOIAS"];

// Nome por extenso pra exibição (card da Unidade + breadcrumb) — o dado
// bruto (`r.Unidade`, usado em filtros/tabela/data-valor) continua a sigla.
const NOME_COMPLETO_UNIDADE = {
  SP: "São Paulo",
  RJ: "Rio de Janeiro",
  Santos: "Santos",
  GOIAS: "Goiás",
};

function nomeCompletoUnidade(sigla) {
  return NOME_COMPLETO_UNIDADE[sigla] || sigla;
}

function renderizarCardsUnidades(container, rows, aoClicar) {
  const unidades = [...new Set(rows.map((r) => r.Unidade).filter(Boolean))].sort((a, b) => {
    const ia = ORDEM_UNIDADES.indexOf(a);
    const ib = ORDEM_UNIDADES.indexOf(b);
    if (ia === -1 && ib === -1) return a.localeCompare(b, "pt-BR");
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });
  if (!unidades.length) {
    container.innerHTML = `<p class="evolucao-vazio">Nenhuma unidade com dados para este tipo de relatório.</p>`;
    return;
  }
  container.innerHTML = unidades
    .map((nome) => `
      <div class="unidade-card" data-valor="${nome.replace(/"/g, "&quot;")}">
        <div class="unidade-card-nome">${nomeCompletoUnidade(nome)}</div>
        <div class="unidade-card-footer">Clique para ver os detalhes</div>
      </div>
    `)
    .join("");

  container.querySelectorAll(".unidade-card").forEach((cardEl) => {
    cardEl.addEventListener("click", () => aoClicar(cardEl.dataset.valor));
  });
}

function renderizarQuebraGrupo(container, campo, filtroEl) {
  const grupos = contarDetalhado(filtrados, campo);
  const selecionado = filtroEl.value;
  container.innerHTML = grupos
    .map(([nome, g]) => {
      const ativo = nome === selecionado ? " selecionado" : "";
      const docsHtml = ORDEM_DOCUMENTACAO
        .map((docNome) => renderizarDocGrupo(docNome, g.docs.get(docNome), g.total))
        .join("");

      return `
        <div class="quebra-card${ativo}" data-campo="${campo}" data-valor="${nome.replace(/"/g, "&quot;")}" title="${nome}">
          <div class="quebra-cabecalho">
            <div class="quebra-nome">${nome}</div>
            <div class="quebra-total-num">${g.total.toLocaleString("pt-BR")}</div>
          </div>
          <div class="quebra-docs">${docsHtml}</div>
        </div>
      `;
    })
    .join("");

  container.querySelectorAll(".quebra-card").forEach((cardEl) => {
    cardEl.addEventListener("click", () => alternarFiltroEMostrarTabela(cardEl.dataset.campo, cardEl.dataset.valor));
  });
}

function renderizarQuebras() {
  renderizarQuebraGrupo(el.quebraConteudo, "Tributacao", el.regime);
}

function renderizarRankingGerentes() {
  const contagens = new Map();
  filtrados.forEach((r) => {
    const gerente = r.Gerente;
    if (!gerente) return;
    if (!contagens.has(gerente)) contagens.set(gerente, { total: 0, pendente: 0 });
    const c = contagens.get(gerente);
    c.total++;
    if (r.Documentacao !== "Documentação Recebida") c.pendente++;
  });

  const lista = [...contagens.entries()]
    .filter(([, c]) => c.pendente > 0)
    .sort((a, b) => b[1].pendente - a[1].pendente)
    .slice(0, 10);

  if (!lista.length) {
    el.rankingGerentes.innerHTML = `<p style="color:var(--cinza-muted); font-size:0.85rem; margin:8px 0 0;">Nenhuma pendência no filtro atual.</p>`;
    return;
  }

  const maior = Math.max(...lista.map(([, c]) => c.pendente));
  const selecionado = el.gerente.value;

  el.rankingGerentes.innerHTML = lista
    .map(([nome, c]) => {
      const largura = (c.pendente / maior) * 100;
      const ativo = nome === selecionado ? " selecionado" : "";
      return `
        <div class="ranking-linha${ativo}" data-valor="${nome.replace(/"/g, "&quot;")}" title="${nome}: ${c.pendente} de ${c.total} pendente(s)">
          <div class="ranking-rotulo">${nome}</div>
          <div class="ranking-trilha"><div class="ranking-barra" style="width:${largura}%"></div></div>
          <div class="ranking-valor">${c.pendente}</div>
        </div>
      `;
    })
    .join("");

  // O número mostrado no ranking é só as pendências do gerente (c.pendente),
  // não o total de clientes dele — o filtro precisa combinar Gerente +
  // Documentação Pendente pra mostrar exatamente esse número na tabela
  // (mesma correção de filtrarPorVariosEMostrarTabela acima).
  el.rankingGerentes.querySelectorAll(".ranking-linha").forEach((linhaEl) => {
    linhaEl.addEventListener("click", () => filtrarPorVariosEMostrarTabela([
      ["Gerente", linhaEl.dataset.valor],
      ["Documentacao", "Documentação Pendente"],
    ]));
  });
}

function celula(texto) {
  return texto === null || texto === undefined || texto === "" ? "—" : texto;
}

function nomeComId(id, nome) {
  const rotuloNome = celula(nome);
  return id === null || id === undefined || id === "" ? rotuloNome : `${id} - ${rotuloNome}`;
}

// Defesa: o backend de cada fonte já remove o prefixo "Federal -" de
// Tributacao (MAPA_REGIME/MAPA_TRIBUTACAO), este é só um fallback caso um
// dia isso mude.
function regimeCurto(texto) {
  if (!texto) return "—";
  return texto.replace(/^Federal\s*-\s*/, "");
}

function renderizarTabela() {
  el.corpo.innerHTML = filtradosTabela
    .map((r) => {
      const doc = r.Documentacao;
      const rotuloDoc = doc ? doc.replace("Documentação ", "") : "—";
      return `
        <tr>
          <td>${nomeComId(r.Id, r.Cliente)}</td>
          <td>${celula(r.Grupo)}</td>
          <td>${celula(r.Unidade ? r.Unidade.toUpperCase() : r.Unidade)}</td>
          <td>${celula(r.Segmento)}</td>
          <td>${celula(r.Gerente)}</td>
          <td>${celula(r.Departamento)}</td>
          <td>${regimeCurto(r.Tributacao)}</td>
          <td>${celula(rotuloStatus(r.Status))}</td>
          <td>${rotuloDoc}</td>
          <td title="${r.DocumentoPendente ? r.DocumentoPendente.replace(/"/g, "&quot;") : ""}">${celula(r.DocumentoPendente)}</td>
        </tr>
      `;
    })
    .join("");

  el.contagem.textContent = `${filtradosTabela.length.toLocaleString("pt-BR")} empresa(s)`;
}

// Cada fonte tem seu próprio status.json (última execução do robô
// correspondente) — guardados aqui pra trocar o texto exibido no header/
// rodapé conforme o Tipo de Relatório ativo, sem precisar buscar de novo.
const STATUS_URL_POR_TIPO = {
  "Radar Fiscal": "data/relatorio_fechamentos/status_radar_fiscal.json",
  "Análise de Balanço": "data/relatorio_fechamentos/status_analise_balanco.json",
};
const statusPorTipo = {};

function carregarStatus() {
  const chaves = Object.keys(STATUS_URL_POR_TIPO);
  Promise.all(
    chaves.map((tipo) =>
      fetch(STATUS_URL_POR_TIPO[tipo] + "?" + Date.now())
        .then((r) => r.json())
        .then((s) => { statusPorTipo[tipo] = s; })
        .catch(() => { statusPorTipo[tipo] = null; })
    )
  ).then(atualizarStatusExibido);
}

function atualizarStatusExibido() {
  const s = statusPorTipo[tipoRelatorioAtivo];
  if (!s) {
    el.status.textContent = "Nenhuma execução registrada ainda.";
    return;
  }
  const data = new Date(s.ultima_execucao);
  el.status.textContent = `${tipoRelatorioAtivo} atualizado em ${data.toLocaleString("pt-BR")}`;
}

function formatarDataCurta(iso) {
  return new Date(`${iso}T00:00:00`).toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });
}

// Quantas empresas foram fechadas (DataReferencia — DataConfirmacao no
// Radar Fiscal, DataImportacao na Análise de Balanço) em cada dia — vem
// direto da planilha (cada linha já tem sua própria data), não de um
// histórico acumulado por execução do robô. Respeita os filtros ativos
// (mesmo conjunto `filtrados` dos cards/ranking).
function contarConfirmacoesPorDia() {
  const dias = new Map();
  filtrados.forEach((r) => {
    if (!r.DataReferencia) return;
    const dia = r.DataReferencia.slice(0, 10);
    dias.set(dia, (dias.get(dia) || 0) + 1);
  });
  return [...dias.entries()]
    .map(([data, total]) => ({ data, total }))
    .sort((a, b) => a.data.localeCompare(b.data));
}

// Uma barra por dia (SVG desenhado à mão, sem lib externa — mesmo padrão do
// resto do portal) com o total de empresas fechadas naquele dia. Barra (não
// linha) porque é uma contagem discreta por dia, não um total acumulado.
function renderizarEvolucao() {
  const container = el.evolucaoGrafico;
  if (!container) return;

  const historico = contarConfirmacoesPorDia();
  if (historico.length < 2) {
    container.innerHTML = `<p class="evolucao-vazio">Sem dias suficientes com data de referência no filtro atual para montar o gráfico.</p>`;
    return;
  }

  const largura = Math.max(360, Math.round(container.clientWidth || 900));
  const altura = 380;
  const margemEsq = 46, margemDir = 16, margemTopo = 26, margemBaixo = 34;
  const areaLargura = largura - margemEsq - margemDir;
  const areaAltura = altura - margemTopo - margemBaixo;

  const maiorTotal = Math.max(1, ...historico.map((h) => h.total));
  const passoX = areaLargura / historico.length;
  const larguraBarra = passoX * 0.6;

  const coordXCentro = (i) => margemEsq + i * passoX + passoX / 2;
  const alturaBarra = (valor) => (valor / maiorTotal) * areaAltura;
  const coordY = (valor) => margemTopo + areaAltura - (valor / maiorTotal) * areaAltura;

  const barras = historico
    .map((h, i) => {
      const x = coordXCentro(i) - larguraBarra / 2;
      const yBase = margemTopo + areaAltura;
      const alt = alturaBarra(h.total);
      const y = yBase - alt;
      const rotuloTotal = `<text x="${coordXCentro(i)}" y="${y - 6}" text-anchor="middle" class="evolucao-rotulo-total">${h.total.toLocaleString("pt-BR")}</text>`;
      return `
        <rect x="${x}" y="${y}" width="${larguraBarra}" height="${alt}" class="evolucao-barra">
          <title>${formatarDataCurta(h.data)} — ${h.total.toLocaleString("pt-BR")} fechado(s)</title>
        </rect>
        ${rotuloTotal}
      `;
    })
    .join("");

  const NUM_GRADES = 4;
  const grades = Array.from({ length: NUM_GRADES + 1 }, (_, i) => {
    const valor = Math.round((maiorTotal / NUM_GRADES) * i);
    const y = coordY(valor);
    return `
      <line x1="${margemEsq}" y1="${y}" x2="${largura - margemDir}" y2="${y}" class="evolucao-grade" />
      <text x="${margemEsq - 8}" y="${y + 4}" text-anchor="end" class="evolucao-eixo-texto">${valor.toLocaleString("pt-BR")}</text>
    `;
  }).join("");

  const rotulosX = historico
    .map((h, i) => `<text x="${coordXCentro(i)}" y="${altura - margemBaixo + 18}" text-anchor="middle" class="evolucao-eixo-texto">${formatarDataCurta(h.data)}</text>`)
    .join("");

  container.innerHTML = `
    <svg viewBox="0 0 ${largura} ${altura}" class="evolucao-svg" role="img" aria-label="Fechamentos por dia">
      ${grades}
      ${barras}
      ${rotulosX}
    </svg>
  `;
}

const UNIDADES_EXCLUIDAS = ["MG EXPRESS"];

// Ordem fixa das abas do topo (Radar Fiscal primeiro, igual à ordem que o
// usuário pediu pra juntar os dois portais).
const ORDEM_TIPO_RELATORIO = ["Radar Fiscal", "Análise de Balanço"];

// A "2ª dimensão" de navegação é "Departamento" no Radar Fiscal (aninhado
// só nele) e "Segmento" na Análise de Balanço (que não tem Departamento) —
// mesmo papel estrutural, fonte de dado diferente. Usado pela tela de
// Departamento/Segmento (ver renderizarTelaDepartamentos).
const QUEBRA_CONFIG_POR_TIPO = {
  "Radar Fiscal": { segunda: { chave: "Departamento", label: "Por Departamento" } },
  "Análise de Balanço": { segunda: { chave: "Segmento", label: "Por Segmento" } },
};

function renderizarTipoRelatorioAbas(tipos) {
  const ordenados = [...tipos].sort((a, b) => {
    const ia = ORDEM_TIPO_RELATORIO.indexOf(a);
    const ib = ORDEM_TIPO_RELATORIO.indexOf(b);
    if (ia === -1 && ib === -1) return a.localeCompare(b, "pt-BR");
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });

  el.tipoRelatorioAbas.innerHTML = ordenados
    .map((tipo) => `<button type="button" class="tipo-relatorio-aba" data-valor="${tipo.replace(/"/g, "&quot;")}" role="tab">${tipo}</button>`)
    .join("");

  el.tipoRelatorioAbas.querySelectorAll(".tipo-relatorio-aba").forEach((botao) => {
    botao.addEventListener("click", () => {
      if (botao.dataset.valor !== tipoRelatorioAtivo) selecionarTipoRelatorio(botao.dataset.valor);
    });
  });

  return ordenados;
}

// ── Navegação Unidade -> Departamento/Segmento ──────────────────────────
// `escopo.unidade` null = tela de Unidades. `escopo.unidade` setado e
// `escopo.depto` null = tela da Unidade (visão Consolidada da unidade
// inteira + cards de Departamento/Segmento pra detalhar — as "2 visões").
// Os dois setados = tela do Departamento/Segmento (corpo do dashboard mais
// restrito).
function irParaTelaUnidades() {
  escopo = { unidade: null, depto: null };
  atualizarNavegacao();
}

function irParaUnidadeConsolidado() {
  escopo.depto = null;
  atualizarNavegacao();
}

function selecionarUnidade(unidade) {
  escopo = { unidade, depto: null };
  atualizarNavegacao();
}

function selecionarDepartamento(depto) {
  escopo.depto = depto;
  atualizarNavegacao();
}

// Botão "Voltar" — volta 1 nível por vez (departamento -> unidade -> painel
// de unidades), não pula direto pro painel de controle igual o crumb
// "Painel de Unidades" do breadcrumb.
function voltarUmaSecao() {
  if (escopo.depto) irParaUnidadeConsolidado();
  else irParaTelaUnidades();
}

function renderizarBreadcrumb() {
  const partes = [{ texto: "Painel de Unidades", acao: irParaTelaUnidades }];
  if (escopo.unidade) partes.push({ texto: nomeCompletoUnidade(escopo.unidade), acao: irParaUnidadeConsolidado });
  if (escopo.depto) partes.push({ texto: escopo.depto, acao: null });

  el.breadcrumbCrumbs.innerHTML = partes
    .map((p, i) => {
      if (i === partes.length - 1) return `<span class="crumb-atual">${p.texto}</span>`;
      return `<span class="crumb-link" data-i="${i}">${p.texto}</span><span class="crumb-sep">›</span>`;
    })
    .join("");

  el.breadcrumbCrumbs.querySelectorAll(".crumb-link").forEach((crumbEl) => {
    crumbEl.addEventListener("click", () => partes[Number(crumbEl.dataset.i)].acao());
  });

  el.breadcrumbBar.classList.toggle("oculto", !escopo.unidade);
}

function calcularDadosEscopo() {
  const chaveSegunda = QUEBRA_CONFIG_POR_TIPO[tipoRelatorioAtivo].segunda.chave;
  return dadosTipo.filter((r) => {
    if (escopo.unidade && r.Unidade !== escopo.unidade) return false;
    if (escopo.depto && r[chaveSegunda] !== escopo.depto) return false;
    return true;
  });
}

function limparFiltrosGerais() {
  el.busca.value = "";
  el.segmento.value = "";
  el.regime.value = "";
  el.status_.value = "";
  el.documentacao.value = "";
  el.gerente.value = "";
}

function limparFiltrosTabela() {
  el.tBusca.value = "";
  el.tSegmento.value = "";
  el.tRegime.value = "";
  el.tStatus.value = "";
  el.tDocumentacao.value = "";
  el.tGerente.value = "";
}

// Corpo do dashboard (Filtros, Por Tributação, Evolução, Ranking, Filtros
// da tabela, Tabela) — compartilhado entre a tela da Unidade (Consolidado)
// e a tela do Departamento/Segmento, só muda o recorte de `dadosEscopo`.
function atualizarCorpoDashboard() {
  dadosEscopo = calcularDadosEscopo();

  limparFiltrosGerais();
  limparFiltrosTabela();

  repopularSelect(el.segmento, new Set(dadosEscopo.map((r) => r.Segmento).filter(Boolean)));
  repopularSelect(el.regime, new Set(dadosEscopo.map((r) => r.Tributacao).filter(Boolean)));
  repopularSelect(el.gerente, new Set(dadosEscopo.map((r) => r.Gerente).filter(Boolean)));
  repopularSelect(el.status_, new Set(dadosEscopo.map((r) => r.Status).filter(Boolean)), rotuloStatus);
  repopularSelect(el.documentacao, new Set(dadosEscopo.map((r) => r.Documentacao).filter(Boolean)));
  repopularSelect(el.tSegmento, new Set(dadosEscopo.map((r) => r.Segmento).filter(Boolean)));
  repopularSelect(el.tRegime, new Set(dadosEscopo.map((r) => r.Tributacao).filter(Boolean)));
  repopularSelect(el.tGerente, new Set(dadosEscopo.map((r) => r.Gerente).filter(Boolean)));
  repopularSelect(el.tStatus, new Set(dadosEscopo.map((r) => r.Status).filter(Boolean)), rotuloStatus);
  repopularSelect(el.tDocumentacao, new Set(dadosEscopo.map((r) => r.Documentacao).filter(Boolean)));

  aplicarFiltros();
  aplicarFiltroTabela();
}

function atualizarNavegacao() {
  renderizarBreadcrumb();

  const telaUnidades = !escopo.unidade;
  const telaDepartamentoGrid = !!escopo.unidade && !escopo.depto;
  const corpoVisivel = !!escopo.unidade;

  el.secaoUnidades.classList.toggle("oculto", !telaUnidades);
  el.secaoDepartamentos.classList.toggle("oculto", !telaDepartamentoGrid);
  el.corpoDashboard.classList.toggle("oculto", !corpoVisivel);

  if (telaUnidades) {
    renderizarCardsUnidades(el.unidadesGrid, dadosTipo, selecionarUnidade);
    return;
  }

  if (telaDepartamentoGrid) {
    const cfgSegunda = QUEBRA_CONFIG_POR_TIPO[tipoRelatorioAtivo].segunda;
    el.departamentosGridTitulo.textContent = cfgSegunda.label;
    const rowsUnidade = dadosTipo.filter((r) => r.Unidade === escopo.unidade);
    renderizarCardsNavegacao(
      el.departamentosGrid, rowsUnidade, cfgSegunda.chave, selecionarDepartamento,
      "Nenhum registro para esta unidade."
    );
  }

  atualizarCorpoDashboard();
}

function selecionarTipoRelatorio(tipo) {
  tipoRelatorioAtivo = tipo;
  dadosTipo = dados.filter((r) => r.TipoRelatorio === tipo);

  el.tipoRelatorioAbas.querySelectorAll(".tipo-relatorio-aba").forEach((botao) => {
    const ativo = botao.dataset.valor === tipo;
    botao.classList.toggle("ativa", ativo);
    botao.setAttribute("aria-selected", ativo ? "true" : "false");
  });

  // Trocar de Tipo de Relatório volta pra tela de Unidades — mesmo padrão
  // de sempre resetar ao trocar de fonte, evita manter selecionada uma
  // Unidade/Departamento que pode não existir na outra fonte.
  escopo = { unidade: null, depto: null };

  atualizarStatusExibido();
  atualizarNavegacao();
}

function carregarDados() {
  Promise.all([
    fetch("data/relatorio_fechamentos/radar_fiscal_dados.json?" + Date.now())
      .then((r) => r.json())
      .then((json) => json.map(normalizarRadarFiscal))
      .catch(() => []),
    fetch("data/relatorio_fechamentos/analise_balanco_dados.json?" + Date.now())
      .then((r) => r.json())
      .then((json) => json.map(normalizarAnaliseBalanco))
      .catch(() => []),
  ]).then(([radarFiscal, analiseBalanco]) => {
    dados = [...radarFiscal, ...analiseBalanco].filter((r) => !UNIDADES_EXCLUIDAS.includes(r.Unidade));
    if (!dados.length) {
      el.unidadesGrid.innerHTML = `<p class="evolucao-vazio">Nenhum dado exportado ainda — rode os robôs do Radar Fiscal e da Análise de Balanço.</p>`;
      return;
    }
    const tipos = renderizarTipoRelatorioAbas([...new Set(dados.map((r) => r.TipoRelatorio).filter(Boolean))]);
    selecionarTipoRelatorio(tipos[0] || null);
  });
}

[el.busca, el.segmento, el.regime, el.status_, el.documentacao, el.gerente].forEach((campo) => {
  campo.addEventListener("input", aplicarFiltros);
  campo.addEventListener("change", aplicarFiltros);
});

el.limpar.addEventListener("click", () => {
  limparFiltrosGerais();
  aplicarFiltros();
});

[el.tBusca, el.tSegmento, el.tRegime, el.tStatus, el.tDocumentacao, el.tGerente].forEach((campo) => {
  campo.addEventListener("input", aplicarFiltroTabela);
  campo.addEventListener("change", aplicarFiltroTabela);
});

el.tLimpar.addEventListener("click", () => {
  limparFiltrosTabela();
  aplicarFiltroTabela();
});

const elBtnTema = document.getElementById("btn-tema");
elBtnTema.addEventListener("click", () => {
  const escuro = document.body.classList.toggle("tema-escuro");
  elBtnTema.textContent = escuro ? "Alterar tema para Claro" : "Alterar tema para Escuro";
});

el.btnVoltarPainel.addEventListener("click", voltarUmaSecao);

// Re-renderiza "Evolução Diária" quando a largura do container muda (ex.:
// redimensionar a janela) — o viewBox do gráfico é calculado a partir da
// largura real do container (ver renderizarEvolucao), então precisa
// recalcular pra manter 1 unidade do SVG = 1px real.
let resizeTimeout;
window.addEventListener("resize", () => {
  clearTimeout(resizeTimeout);
  resizeTimeout = setTimeout(renderizarEvolucao, 200);
});

carregarStatus();
carregarDados();
