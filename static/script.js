// Portal único que junta o Radar Fiscal e a Análise de Balanço (mesma ideia
// da Análise de Entrega de SPED com ICMS/Contribuições): os dois relatórios
// têm robôs e planilhas totalmente diferentes, então cada fonte é
// normalizada pra um esquema comum (ver normalizarRadarFiscal/
// normalizarAnaliseBalanco) antes de entrar em `dados`. A partir daí, a
// página inteira (filtros, cards, ranking, evolução, tabela) sempre mostra
// só um Tipo de Relatório por vez, escolhido na aba do topo — igual ao
// padrão do SPED.
let dados = [];
let dadosTipo = [];
let tipoRelatorioAtivo = null;
let filtrados = [];
let filtradosTabela = [];

const el = {
  status: document.getElementById("status-execucao"),
  statusRodape: document.getElementById("status-execucao-rodape"),
  tipoRelatorioAbas: document.getElementById("tipo-relatorio-abas"),
  busca: document.getElementById("f-busca"),
  // Não é um <select> — é um estado simples com um Set de valores (permite
  // marcar mais de uma unidade ao mesmo tempo). filtrarConjunto trata esse
  // campo de forma diferente do resto (ver `.valores` lá). Os botões ficam
  // soltos no topo da página (unidadeTopoLista).
  unidade: { valores: new Set() },
  unidadeTopoLista: document.getElementById("unidade-topo-lista"),
  segmento: document.getElementById("f-segmento"),
  regime: document.getElementById("f-regime"),
  depto: document.getElementById("f-depto"),
  deptoGrupo: document.getElementById("f-depto-grupo"),
  status_: document.getElementById("f-status"),
  documentacao: document.getElementById("f-documentacao"),
  gerente: document.getElementById("f-gerente"),
  limpar: document.getElementById("f-limpar"),
  corpo: document.getElementById("tabela-corpo"),
  contagem: document.getElementById("contagem"),
  quebraConteudo: document.getElementById("quebra-conteudo"),
  quebraAbas: document.querySelectorAll("#quebra-abas-dimensao .quebra-aba"),
  quebraAbaSegunda: document.getElementById("quebra-aba-segunda"),
  rankingGerentes: document.getElementById("ranking-gerentes"),
  evolucaoGrafico: document.getElementById("evolucao-grafico"),
  // Filtro independente, só da tabela — não afeta KPIs/cards/ranking
  tBusca: document.getElementById("t-busca"),
  tUnidade: document.getElementById("t-unidade"),
  tSegmento: document.getElementById("t-segmento"),
  tRegime: document.getElementById("t-regime"),
  tDepto: document.getElementById("t-depto"),
  tDeptoGrupo: document.getElementById("t-depto-grupo"),
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
    Documentacao: r["Documentação"],
    Departamento: r.DeptoFiscal,
    DataReferencia: r.DataConfirmacao,
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
    Documentacao: r["Documentação"],
    Departamento: undefined,
    DataReferencia: r.DataImportacao,
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
// trocar de Tipo de Relatório, já que os valores possíveis de cada filtro
// mudam de uma fonte pra outra.
function repopularSelect(select, valores, formatar = (v) => v) {
  const placeholder = select.options[0];
  select.innerHTML = "";
  select.appendChild(placeholder);
  popularSelect(select, valores, formatar);
}

function filtrarConjunto(conjunto, campos) {
  const busca = campos.busca.value.trim().toLowerCase();
  // Unidade aceita dois formatos: um <select> normal (.value, string única
  // — usado no filtro da tabela) ou o estado multi-seleção dos chips do
  // topo (.valores, Set — vazio = "Todas").
  const unidadeValores = campos.unidade.valores;
  const unidadeUnica = campos.unidade.value;
  const segmento = campos.segmento.value;
  const regime = campos.regime.value;
  const depto = campos.depto.value;
  const status = campos.status.value;
  const documentacao = campos.documentacao.value;
  const gerente = campos.gerente.value;

  return conjunto.filter((r) => {
    if (unidadeValores) {
      if (unidadeValores.size > 0 && !unidadeValores.has(r.Unidade)) return false;
    } else if (unidadeUnica && r.Unidade !== unidadeUnica) {
      return false;
    }
    if (segmento && r.Segmento !== segmento) return false;
    if (regime && r.Tributacao !== regime) return false;
    if (depto && r.Departamento !== depto) return false;
    if (gerente && r.Gerente !== gerente) return false;
    if (status && r.Status !== status) return false;
    if (documentacao && r.Documentacao !== documentacao) return false;
    if (busca) {
      const alvo = `${r.Cliente || ""} ${r.Grupo || ""}`.toLowerCase();
      if (!alvo.includes(busca)) return false;
    }
    return true;
  });
}

function aplicarFiltros() {
  filtrados = filtrarConjunto(dadosTipo, {
    busca: el.busca, unidade: el.unidade, segmento: el.segmento, regime: el.regime,
    depto: el.depto, status: el.status_, documentacao: el.documentacao, gerente: el.gerente,
  });

  renderizarQuebras();
  renderizarRankingGerentes();
  renderizarEvolucao();
  atualizarUnidadeTopoAtiva();
}

function renderizarUnidadeTopo() {
  const unidades = [...new Set(dadosTipo.map((r) => r.Unidade).filter(Boolean))].sort((a, b) => a.localeCompare(b, "pt-BR"));
  el.unidadeTopoLista.innerHTML = [`<button type="button" class="unidade-topo-chip" data-valor="">Todas</button>`]
    .concat(
      unidades.map((u) => `<button type="button" class="unidade-topo-chip" data-valor="${u.replace(/"/g, "&quot;")}">${u.toUpperCase()}</button>`)
    )
    .join("");
  ativarFiltroUnidadeMultiplo();
  atualizarUnidadeTopoAtiva();
}

// Diferente de ativarFiltroClicavel (usado por Tributação/Departamento-ou-
// Segmento/Gerente, que são single-select): aqui cada clique liga/desliga
// aquela unidade sem desmarcar as outras, permitindo comparar várias de uma
// vez. "Todas" é um caso especial que limpa a seleção inteira.
function ativarFiltroUnidadeMultiplo() {
  el.unidadeTopoLista.querySelectorAll(".unidade-topo-chip").forEach((botao) => {
    botao.addEventListener("click", () => {
      const valor = botao.dataset.valor;
      if (valor === "") {
        el.unidade.valores.clear();
      } else if (el.unidade.valores.has(valor)) {
        el.unidade.valores.delete(valor);
      } else {
        el.unidade.valores.add(valor);
      }
      aplicarFiltros();
    });
  });
}

function atualizarUnidadeTopoAtiva() {
  el.unidadeTopoLista.querySelectorAll(".unidade-topo-chip").forEach((botao) => {
    const valor = botao.dataset.valor;
    const ativo = valor === "" ? el.unidade.valores.size === 0 : el.unidade.valores.has(valor);
    botao.classList.toggle("ativo", ativo);
  });
}

function aplicarFiltroTabela() {
  filtradosTabela = filtrarConjunto(dadosTipo, {
    busca: el.tBusca, unidade: el.tUnidade, segmento: el.tSegmento, regime: el.tRegime,
    depto: el.tDepto, status: el.tStatus, documentacao: el.tDocumentacao, gerente: el.tGerente,
  });
  renderizarTabela();
}

const ORDEM_DOCUMENTACAO = ["Documentação Recebida", "Documentação Pendente"];
// Os 5 valores reais de Status são diferentes entre as duas fontes — cada
// uma tem uma categoria sem equivalente na outra (Radar Fiscal: "Bloqueado";
// Análise de Balanço: "Importado Contábil"). Por isso a ordem/lista é por
// Tipo de Relatório, não global — um card do Radar Fiscal nunca mostra uma
// linha fixa de "Importado Contábil: 0" que não faz sentido pra ele.
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

function contarDetalhado(chave) {
  const grupos = new Map();
  filtrados.forEach((r) => {
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

function ativarFiltroClicavel(elementos, filtroEl) {
  elementos.forEach((el2) => {
    el2.addEventListener("click", () => {
      const valor = el2.dataset.valor;
      filtroEl.value = filtroEl.value === valor ? "" : valor;
      aplicarFiltros();
    });
  });
}

function renderizarDocGrupo(docNome, d, totalCategoria) {
  const classe = docNome === "Documentação Recebida" ? "recebida" : "pendente";
  const pctDoc = totalCategoria ? (d.total / totalCategoria) * 100 : 0;
  // Ordem fixa (não por contagem) pra todo card mostrar as mesmas linhas na
  // mesma posição — mesmo as zeradas.
  const statusOrdenado = statusOrdem().map((s) => [s, d.status.get(s) || 0]);

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

function renderizarQuebraGrupo(container, chave, filtroEl) {
  const grupos = contarDetalhado(chave);
  const selecionado = filtroEl.value;
  container.innerHTML = grupos
    .map(([nome, g]) => {
      const ativo = nome === selecionado ? " selecionado" : "";
      const docsHtml = ORDEM_DOCUMENTACAO
        .map((docNome) => renderizarDocGrupo(docNome, g.docs.get(docNome), g.total))
        .join("");

      return `
        <div class="quebra-card${ativo}" data-valor="${nome.replace(/"/g, "&quot;")}" title="${nome}">
          <div class="quebra-cabecalho">
            <div class="quebra-nome">${nome}</div>
            <div class="quebra-total-num">${g.total.toLocaleString("pt-BR")}</div>
          </div>
          <div class="quebra-docs">${docsHtml}</div>
        </div>
      `;
    })
    .join("");

  ativarFiltroClicavel(container.querySelectorAll(".quebra-card"), filtroEl);
}

function contarDetalhadoComSubgrupo(chavePrincipal, chaveSecundaria) {
  const grupos = new Map();
  filtrados.forEach((r) => {
    const valor1 = r[chavePrincipal];
    if (!valor1) return;
    if (!grupos.has(valor1)) grupos.set(valor1, { total: 0, sub: new Map() });
    const g = grupos.get(valor1);
    g.total++;

    const valor2 = r[chaveSecundaria] || "Sem tributação";
    if (!g.sub.has(valor2)) g.sub.set(valor2, { total: 0, docs: criarContadorDocs() });
    const s = g.sub.get(valor2);
    s.total++;

    const doc = r.Documentacao || "Sem documentação";
    if (!s.docs.has(doc)) s.docs.set(doc, criarContadorStatus());
    const d = s.docs.get(doc);
    d.total++;

    const status = r.Status || "Não importado";
    d.status.set(status, (d.status.get(status) || 0) + 1);
  });
  return [...grupos.entries()].sort((a, b) => b[1].total - a[1].total);
}

function renderizarRegimeCard(nome, s, selecionado) {
  const ativo = nome === selecionado ? " selecionado" : "";
  const docsHtml = ORDEM_DOCUMENTACAO
    .map((docNome) => renderizarDocGrupo(docNome, s.docs.get(docNome), s.total))
    .join("");

  return `
    <div class="regime-card${ativo}" data-valor="${nome.replace(/"/g, "&quot;")}" title="${nome}">
      <div class="regime-card-cabecalho">
        <div class="regime-card-nome">${nome}</div>
        <div class="regime-card-total">${s.total.toLocaleString("pt-BR")}</div>
      </div>
      <div class="quebra-docs">${docsHtml}</div>
    </div>
  `;
}

const faixasAbertas = new Set();

// Clicar no cabeçalho da 2ª dimensão (Departamento no Radar Fiscal, Segmento
// na Análise de Balanço) só expande/recolhe (não filtra) — pra filtrar, usa
// o botão "Fixar". Card de Tributação dentro continua filtrando (el.regime)
// num clique direto, igual antes.
function renderizarFaixaSegunda(container, chavePrincipal, chaveSecundaria, filtroPrincipal, filtroSecundario) {
  const grupos = contarDetalhadoComSubgrupo(chavePrincipal, chaveSecundaria);
  const selecionado = filtroPrincipal.value;
  const selecionadoSub = filtroSecundario.value;
  container.innerHTML = grupos
    .map(([nome, g]) => {
      const ativo = nome === selecionado ? " selecionado" : "";
      const aberto = faixasAbertas.has(nome);
      const cardsHtml = [...g.sub.entries()]
        .sort((a, b) => b[1].total - a[1].total)
        .map(([subNome, s]) => renderizarRegimeCard(subNome, s, selecionadoSub))
        .join("");

      return `
        <div class="quebra-faixa${aberto ? "" : " colapsado"}${ativo}" title="${nome}">
          <div class="quebra-faixa-cabecalho" data-valor="${nome.replace(/"/g, "&quot;")}">
            <button type="button" class="quebra-faixa-toggle" aria-label="Mostrar detalhe" aria-expanded="${aberto}"><i class="seta"></i></button>
            <div class="quebra-nome">${nome}</div>
            <div class="quebra-total-num">${g.total.toLocaleString("pt-BR")}</div>
            <button type="button" class="quebra-faixa-fixar${ativo ? " fixado" : ""}" aria-label="${ativo ? "Remover filtro" : "Filtrar por este item"}" title="${ativo ? "Remover filtro" : "Fixar filtro"}">
              <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M14.5 2.5a1 1 0 0 1 1.42 0l5.58 5.58a1 1 0 0 1 0 1.42l-1.3 1.3a1 1 0 0 1-1.3.1l-.5-.36-3.02 3.02.6 2.98a1 1 0 0 1-.27.92l-1.1 1.1a1 1 0 0 1-1.42 0l-3.4-3.4-5.02 5.02a1 1 0 0 1-1.42-1.42l5.02-5.02-3.4-3.4a1 1 0 0 1 0-1.42l1.1-1.1a1 1 0 0 1 .92-.27l2.98.6 3.02-3.02-.36-.5a1 1 0 0 1 .1-1.3z"/></svg>
            </button>
          </div>
          <div class="regime-cards">${cardsHtml}</div>
        </div>
      `;
    })
    .join("");

  function alternarAberto(cabecalho) {
    const faixa = cabecalho.closest(".quebra-faixa");
    const botaoToggle = cabecalho.querySelector(".quebra-faixa-toggle");
    const colapsado = faixa.classList.toggle("colapsado");
    const nome = cabecalho.dataset.valor;
    if (colapsado) faixasAbertas.delete(nome); else faixasAbertas.add(nome);
    botaoToggle.setAttribute("aria-expanded", String(!colapsado));
    botaoToggle.setAttribute("aria-label", colapsado ? "Mostrar detalhe" : "Ocultar detalhe");
  }

  container.querySelectorAll(".quebra-faixa-toggle").forEach((botao) => {
    botao.addEventListener("click", (e) => {
      e.stopPropagation();
      alternarAberto(botao.closest(".quebra-faixa-cabecalho"));
    });
  });

  container.querySelectorAll(".quebra-faixa-cabecalho").forEach((cabecalho) => {
    cabecalho.addEventListener("click", (e) => {
      if (e.target.closest(".quebra-faixa-fixar")) return;
      alternarAberto(cabecalho);
    });
  });

  container.querySelectorAll(".quebra-faixa-fixar").forEach((botao) => {
    botao.addEventListener("click", (e) => {
      e.stopPropagation();
      const valor = botao.closest(".quebra-faixa-cabecalho").dataset.valor;
      filtroPrincipal.value = filtroPrincipal.value === valor ? "" : valor;
      aplicarFiltros();
    });
  });

  ativarFiltroClicavel(container.querySelectorAll(".regime-card"), filtroSecundario);
}

// A 2ª aba de cards é "Por Departamento" no Radar Fiscal e "Por Segmento" na
// Análise de Balanço (esta não tem coluna de departamento) — mesmo papel
// estrutural, fonte de dado diferente. O rótulo da aba (#quebra-aba-segunda)
// e o filtro que ela usa (el.depto vs el.segmento — este último é o mesmo
// <select> do filtro geral "Segmento") trocam ao selecionar o Tipo de
// Relatório, ver selecionarTipoRelatorio().
const QUEBRA_CONFIG_POR_TIPO = {
  "Radar Fiscal": {
    regime: { chave: "Tributacao", filtroEl: () => el.regime },
    segunda: { chave: "Departamento", subChave: "Tributacao", filtroEl: () => el.depto, subFiltroEl: () => el.regime, label: "Por Departamento" },
  },
  "Análise de Balanço": {
    regime: { chave: "Tributacao", filtroEl: () => el.regime },
    segunda: { chave: "Segmento", subChave: "Tributacao", filtroEl: () => el.segmento, subFiltroEl: () => el.regime, label: "Por Segmento" },
  },
};
let abaQuebraAtiva = "regime";

function renderizarQuebras() {
  const cfg = QUEBRA_CONFIG_POR_TIPO[tipoRelatorioAtivo][abaQuebraAtiva];
  el.quebraConteudo.classList.toggle("quebra-grid--faixas", abaQuebraAtiva === "segunda");
  if (cfg.subChave) {
    renderizarFaixaSegunda(el.quebraConteudo, cfg.chave, cfg.subChave, cfg.filtroEl(), cfg.subFiltroEl());
  } else {
    renderizarQuebraGrupo(el.quebraConteudo, cfg.chave, cfg.filtroEl());
  }
}

el.quebraAbas.forEach((botao) => {
  botao.addEventListener("click", () => {
    abaQuebraAtiva = botao.dataset.aba;
    el.quebraAbas.forEach((b) => {
      const ativa = b === botao;
      b.classList.toggle("ativa", ativa);
      b.setAttribute("aria-selected", ativa ? "true" : "false");
    });
    renderizarQuebras();
  });
});

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

  ativarFiltroClicavel(el.rankingGerentes.querySelectorAll(".ranking-linha"), el.gerente);
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
    el.statusRodape.textContent = "Nenhuma execução registrada ainda.";
    return;
  }
  const data = new Date(s.ultima_execucao);
  const texto = `${tipoRelatorioAtivo} atualizado em ${data.toLocaleString("pt-BR")}`;
  el.status.textContent = texto;
  el.statusRodape.textContent = texto;
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

function selecionarTipoRelatorio(tipo) {
  tipoRelatorioAtivo = tipo;
  dadosTipo = dados.filter((r) => r.TipoRelatorio === tipo);

  el.tipoRelatorioAbas.querySelectorAll(".tipo-relatorio-aba").forEach((botao) => {
    const ativo = botao.dataset.valor === tipo;
    botao.classList.toggle("ativa", ativo);
    botao.setAttribute("aria-selected", ativo ? "true" : "false");
  });

  // O filtro/coluna "Departamento" só existe no Radar Fiscal — some da tela
  // (não só fica vazio) quando o Tipo ativo é a Análise de Balanço.
  const temDepartamento = tipo === "Radar Fiscal";
  el.deptoGrupo.classList.toggle("oculto", !temDepartamento);
  el.tDeptoGrupo.classList.toggle("oculto", !temDepartamento);

  // Rótulo da 2ª aba de cards muda com a fonte ("Por Departamento" vs
  // "Por Segmento") — ver QUEBRA_CONFIG_POR_TIPO.
  el.quebraAbaSegunda.textContent = QUEBRA_CONFIG_POR_TIPO[tipo].segunda.label;

  // Trocar de Tipo de Relatório reseta todos os filtros (gerais e da
  // tabela) — mesmo padrão do SPED, evita manter selecionado um valor que
  // pode não existir na outra fonte.
  el.busca.value = "";
  el.unidade.valores.clear();
  el.segmento.value = "";
  el.regime.value = "";
  el.depto.value = "";
  el.status_.value = "";
  el.documentacao.value = "";
  el.gerente.value = "";
  el.tBusca.value = "";
  el.tUnidade.value = "";
  el.tSegmento.value = "";
  el.tRegime.value = "";
  el.tDepto.value = "";
  el.tStatus.value = "";
  el.tDocumentacao.value = "";
  el.tGerente.value = "";
  abaQuebraAtiva = "regime";
  el.quebraAbas.forEach((b) => {
    const ativa = b.dataset.aba === "regime";
    b.classList.toggle("ativa", ativa);
    b.setAttribute("aria-selected", ativa ? "true" : "false");
  });

  renderizarUnidadeTopo();
  repopularSelect(el.segmento, new Set(dadosTipo.map((r) => r.Segmento).filter(Boolean)));
  repopularSelect(el.regime, new Set(dadosTipo.map((r) => r.Tributacao).filter(Boolean)));
  repopularSelect(el.depto, new Set(dadosTipo.map((r) => r.Departamento).filter(Boolean)));
  repopularSelect(el.gerente, new Set(dadosTipo.map((r) => r.Gerente).filter(Boolean)));
  repopularSelect(el.status_, new Set(dadosTipo.map((r) => r.Status).filter(Boolean)), rotuloStatus);
  repopularSelect(el.documentacao, new Set(dadosTipo.map((r) => r.Documentacao).filter(Boolean)));
  repopularSelect(el.tUnidade, new Set(dadosTipo.map((r) => r.Unidade).filter(Boolean)), (v) => v.toUpperCase());
  repopularSelect(el.tSegmento, new Set(dadosTipo.map((r) => r.Segmento).filter(Boolean)));
  repopularSelect(el.tRegime, new Set(dadosTipo.map((r) => r.Tributacao).filter(Boolean)));
  repopularSelect(el.tDepto, new Set(dadosTipo.map((r) => r.Departamento).filter(Boolean)));
  repopularSelect(el.tGerente, new Set(dadosTipo.map((r) => r.Gerente).filter(Boolean)));
  repopularSelect(el.tStatus, new Set(dadosTipo.map((r) => r.Status).filter(Boolean)), rotuloStatus);
  repopularSelect(el.tDocumentacao, new Set(dadosTipo.map((r) => r.Documentacao).filter(Boolean)));

  atualizarStatusExibido();
  aplicarFiltros();
  aplicarFiltroTabela();
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
      el.corpo.innerHTML = `<tr><td colspan="9">Nenhum dado exportado ainda — rode os robôs do Radar Fiscal e da Análise de Balanço.</td></tr>`;
      return;
    }
    const tipos = renderizarTipoRelatorioAbas([...new Set(dados.map((r) => r.TipoRelatorio).filter(Boolean))]);
    selecionarTipoRelatorio(tipos[0] || null);
  });
}

[el.busca, el.segmento, el.regime, el.depto, el.status_, el.documentacao, el.gerente].forEach((campo) => {
  campo.addEventListener("input", aplicarFiltros);
  campo.addEventListener("change", aplicarFiltros);
});

el.limpar.addEventListener("click", () => {
  el.busca.value = "";
  el.unidade.valores.clear();
  el.segmento.value = "";
  el.regime.value = "";
  el.depto.value = "";
  el.status_.value = "";
  el.documentacao.value = "";
  el.gerente.value = "";
  aplicarFiltros();
});

[el.tBusca, el.tUnidade, el.tSegmento, el.tRegime, el.tDepto, el.tStatus, el.tDocumentacao, el.tGerente].forEach((campo) => {
  campo.addEventListener("input", aplicarFiltroTabela);
  campo.addEventListener("change", aplicarFiltroTabela);
});

el.tLimpar.addEventListener("click", () => {
  el.tBusca.value = "";
  el.tUnidade.value = "";
  el.tSegmento.value = "";
  el.tRegime.value = "";
  el.tDepto.value = "";
  el.tStatus.value = "";
  el.tDocumentacao.value = "";
  el.tGerente.value = "";
  aplicarFiltroTabela();
});

const elBtnTema = document.getElementById("btn-tema");
elBtnTema.addEventListener("click", () => {
  const escuro = document.body.classList.toggle("tema-escuro");
  elBtnTema.textContent = escuro ? "Alterar tema para Claro" : "Alterar tema para Escuro";
});

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
