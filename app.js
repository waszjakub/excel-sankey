/* Sankey Chart v2: content add-in dla Excela
   Działa jak zwykły wykres: zaznacz dane (Źródło | Cel | Wartość | opcjonalnie Kolor) i wstaw dodatek.
   Formatowanie w panelu bocznym, klik w węzeł otwiera jego ustawienia.
   Każdy wstawiony wykres trzyma własne powiązanie z danymi i własne formatowanie (ustawienia w pliku). */
(function () {
  "use strict";

  const VERSION = "2.1";
  const KEY_DATA = "sankeyBinding";
  const KEY_NODES = "sankeyNodesBinding";
  const KEY_FMT = "sankeyFormat";

  const PALETTES = {
    office: ["#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E", "#4F81BD", "#C0504D"],
    office2013: ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E"],
    mono: ["#6B6B6B"],
    pastel: ["#8FB9DE", "#F4B183", "#A9D18E", "#FFD966", "#C9A0DC", "#9DC3E6", "#F8CBAD", "#C5E0B4"]
  };
  const QUICK_COLORS = ["#6B6B6B", "#00A800", "#B30000", "#156082", "#E97132", "#0F9ED5", "#A02B93", "#FFC000"];

  const DEFAULT_FMT = {
    showTitle: false,
    title: "",
    titleSize: 22,
    titleColor: "#1F4E79",
    showValues: true,
    valuePrefix: "",
    valueSuffix: "",
    decimals: 1,
    numberLocale: "pl",
    palette: "office",
    nodeWidth: 24,
    nodePadding: 30,
    nodeAlign: "left",
    keepOrder: true,
    linkColorBy: "target",
    linkOpacity: 0.45,
    fontSize: 12,
    labelColor: "#222222",
    labelBackground: true,
    background: "#FFFFFF",
    imageScale: 2,
    nodes: {}
  };

  // ---------- Dane przykładowe (Apple FY22) ----------
  const SAMPLE_DATA = [
    ["Źródło", "Cel", "Wartość"],
    ["iPhone", "Products", 205.5],
    ["Mac", "Products", 40.2],
    ["iPad", "Products", 29.3],
    ["Wearables", "Products", 41.2],
    ["Products", "Revenue", 316.2],
    ["Services", "Revenue", 78.2],
    ["Revenue", "Gross profit", 170.9],
    ["Revenue", "Cost of revenue", 223.5],
    ["Gross profit", "Operating profit", 119.5],
    ["Gross profit", "Operating expenses", 51.4],
    ["Operating profit", "Net profit", 99.8],
    ["Operating profit", "Tax", 19.3],
    ["Operating profit", "Other", 0.3],
    ["Operating expenses", "R&D", 26.3],
    ["Operating expenses", "SG&A", 25.1],
    ["Cost of revenue", "Product costs", 201.4],
    ["Cost of revenue", "Service costs", 22.1]
  ];
  const G = "#6B6B6B", P = "#00A800", L = "#B30000";
  const SAMPLE_NODES = {
    iPhone: { color: G, note: "7% Y/Y" }, Mac: { color: G, note: "14% Y/Y" }, iPad: { color: G, note: "-8% Y/Y" },
    Wearables: { color: G, note: "7% Y/Y" }, Products: { color: G, note: "6% Y/Y" },
    Services: { color: G, note: "14% Y/Y", column: 1 }, Revenue: { color: G, note: "8% Y/Y", display: 394.3 },
    "Gross profit": { color: P, note: "12% Y/Y" }, "Cost of revenue": { color: L, note: "5% Y/Y" },
    "Operating profit": { color: P, note: "10% Y/Y" }, "Operating expenses": { color: L, note: "17% Y/Y" },
    "Product costs": { color: L, note: "5% Y/Y" }, "Service costs": { color: L, note: "7% Y/Y" },
    "Net profit": { color: P, note: "5% Y/Y" }, Tax: { color: L, note: "33% Y/Y" }, Other: { color: L },
    "R&D": { color: L, note: "20% Y/Y" }, "SG&A": { color: L, note: "14% Y/Y" }
  };
  const SAMPLE_FMT = {
    showTitle: true, title: "Apple FY22 Income Statement", valuePrefix: "$", valueSuffix: "B",
    numberLocale: "en", decimals: 1, nodePadding: 34, nodes: SAMPLE_NODES
  };

  // ---------- Stan ----------
  let inExcel = false;
  let fmt = clone(DEFAULT_FMT);
  let dataBindingId = null;
  let nodesBindingId = null;
  let handlers = [];
  let model = null; // { graph, depth, skipped, address, nodesAddress }
  let lastSvg = null;
  let changeTimer = null, resizeTimer = null, saveTimer = null, drawTimer = null;
  let selectedNode = null;

  const $ = (id) => document.getElementById(id);

  Office.onReady((info) => {
    inExcel = !!(info && info.host === Office.HostType.Excel);
    wireUi();
    if (!inExcel) {
      // Otwarte w zwykłej przeglądarce: demo, żeby sprawdzić, że hosting działa
      $("btn-copy").hidden = true;
      fmt = Object.assign(clone(DEFAULT_FMT), clone(SAMPLE_FMT));
      model = buildModel(SAMPLE_DATA, null, "demo");
      draw();
      showBanner("Strona dodatku działa (wersja " + VERSION + "). To podgląd demo; w Excelu wykres czyta dane z arkusza.");
      return;
    }
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) $("btn-copy").hidden = true;
    const s = Office.context.document.settings;
    fmt = Object.assign(clone(DEFAULT_FMT), s.get(KEY_FMT) || {});
    dataBindingId = s.get(KEY_DATA) || null;
    nodesBindingId = s.get(KEY_NODES) || null;
    run(async () => {
      if (!dataBindingId) {
        // Pierwsze wstawienie: spróbuj wziąć dane z bieżącego zaznaczenia, jak zwykły wykres
        const ok = await bindFromSelection("data", true);
        if (!ok) {
          showEmpty();
          return;
        }
      }
      await refresh();
      await subscribe();
    });
  });

  // ---------- UI ----------
  function wireUi() {
    $("btn-refresh").addEventListener("click", () => run(inExcel ? refresh : async () => draw()));
    $("btn-data").addEventListener("click", () => run(useSelectionClicked));
    $("btn-format").addEventListener("click", () => togglePane());
    $("btn-copy").addEventListener("click", () => run(insertPngCopy, true));
    $("pane-close").addEventListener("click", () => togglePane(false));
    $("empty-use").addEventListener("click", () => run(useSelectionClicked));
    $("empty-sample").addEventListener("click", () => run(insertSample));

    const pane = $("pane");
    pane.addEventListener("input", onPaneInput);
    pane.addEventListener("change", onPaneInput);
    pane.addEventListener("click", onPaneClick);

    new ResizeObserver(() => {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(() => run(async () => draw()), 100);
    }).observe($("chart"));
  }

  function togglePane(force) {
    const open = typeof force === "boolean" ? force : document.body.classList.contains("pane-open") === false;
    document.body.classList.toggle("pane-open", open);
    if (open) renderPane();
    else if (selectedNode) {
      selectedNode = null;
      scheduleDraw();
    }
  }

  function showEmpty() {
    model = null;
    clearChart();
    $("empty").hidden = false;
    $("empty-use").hidden = !inExcel;
    $("empty-sample").hidden = !inExcel;
  }

  // ---------- Powiązanie z danymi (jak zakres danych wykresu) ----------
  async function useSelectionClicked() {
    const ok = await bindFromSelection("data", false);
    if (!ok) throw new Error("Zaznacz w arkuszu zakres z danymi (Źródło, Cel, Wartość) i kliknij ponownie.");
    selectedNode = null;
    await refresh();
    await subscribe();
    if (document.body.classList.contains("pane-open")) renderPane();
    showBanner("Wykres korzysta teraz z zakresu " + (model ? model.address : "") + ".");
  }

  // kind: "data" | "nodes". silent = true: bez błędu, gdy zaznaczenie nie wygląda na dane
  async function bindFromSelection(kind, silent) {
    try {
      return await bindFromSelectionInner(kind, silent);
    } catch (e) {
      if (silent) return false;
      throw e;
    }
  }

  async function bindFromSelectionInner(kind, silent) {
    let newId = null;
    await Excel.run(async (ctx) => {
      let range = ctx.workbook.getSelectedRange();
      range.load("address,rowCount,columnCount");
      await ctx.sync();

      let tableName = null;
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
        const tables = range.getTables(false);
        tables.load("items/name");
        await ctx.sync();
        if (tables.items.length) tableName = tables.items[0].name;
      }
      if (!tableName && range.rowCount * range.columnCount === 1) {
        // Jedna komórka: weź otaczający blok danych (jak Excel przy wstawianiu wykresu)
        if (Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
          range = range.getSurroundingRegion();
          range.load("address,rowCount,columnCount");
          await ctx.sync();
        }
      }
      if (!tableName && (range.columnCount < (kind === "data" ? 3 : 2) || range.rowCount < 2)) {
        if (silent) return;
        throw new Error(
          kind === "data"
            ? "Zakres musi mieć co najmniej 3 kolumny (Źródło, Cel, Wartość) i 2 wiersze."
            : "Zakres węzłów musi mieć co najmniej 2 kolumny (Węzeł, Kolor/Dopisek...)."
        );
      }

      const oldId = kind === "data" ? dataBindingId : nodesBindingId;
      if (oldId) {
        const old = ctx.workbook.bindings.getItemOrNullObject(oldId);
        await ctx.sync();
        if (!old.isNullObject) old.delete();
      }
      newId = newBindingId(kind);
      await addBinding(ctx, range, tableName, newId);
    });
    if (!newId) return false;
    if (kind === "data") dataBindingId = newId;
    else nodesBindingId = newId;
    $("empty").hidden = true;
    saveSettings(true);
    return true;
  }

  async function unbindNodes() {
    if (!nodesBindingId) return;
    const id = nodesBindingId;
    nodesBindingId = null;
    saveSettings(true);
    await Excel.run(async (ctx) => {
      const b = ctx.workbook.bindings.getItemOrNullObject(id);
      await ctx.sync();
      if (!b.isNullObject) b.delete();
      await ctx.sync();
    });
    await refresh();
    await subscribe();
  }

  function newBindingId(kind) {
    return "sankey" + (kind === "data" ? "Data" : "Nodes") + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  }

  // Powiązanie z tabelą (rośnie z tabelą) jak w przykładach Microsoftu; gdy Excel odmówi, zwykły zakres
  async function addBinding(ctx, range, tableName, id) {
    if (tableName) {
      try {
        ctx.workbook.bindings.add(ctx.workbook.tables.getItem(tableName).getRange(), Excel.BindingType.table, id);
        await ctx.sync();
        return "Table";
      } catch (e) {
        console.warn("Powiązanie typu tabela nie powiodło się, używam zakresu", e);
      }
      range = ctx.workbook.tables.getItem(tableName).getRange();
    }
    ctx.workbook.bindings.add(range, Excel.BindingType.range, id);
    await ctx.sync();
    return "Range";
  }

  async function readBinding(ctx, id) {
    if (!id) return null;
    const b = ctx.workbook.bindings.getItemOrNullObject(id);
    b.load("type");
    await ctx.sync();
    if (b.isNullObject) return null;
    const range = b.type === "Table" ? b.getTable().getRange() : b.getRange();
    range.load("values,address");
    await ctx.sync();
    return { values: range.values, address: range.address };
  }

  async function refresh() {
    let data = null, nodes = null;
    await Excel.run(async (ctx) => {
      data = await readBinding(ctx, dataBindingId);
      nodes = await readBinding(ctx, nodesBindingId);
    });
    if (!data) {
      showEmpty();
      return;
    }
    $("empty").hidden = true;
    model = buildModel(data.values, nodes ? nodes.values : null, data.address, nodes ? nodes.address : null);
    draw();
    if (document.body.classList.contains("pane-open")) renderPane();
    if (model.skipped) showBanner("Pominięto wiersze bez liczby w kolumnie Wartość: " + model.skipped + ".");
  }

  async function subscribe() {
    for (const h of handlers) {
      try {
        await Excel.run(h.context, async (ctx) => {
          h.remove();
          await ctx.sync();
        });
      } catch (e) {
        /* powiązanie mogło zniknąć */
      }
    }
    handlers = [];
    const ids = [dataBindingId, nodesBindingId].filter(Boolean);
    if (!ids.length) return;
    await Excel.run(async (ctx) => {
      const items = ids.map((id) => ctx.workbook.bindings.getItemOrNullObject(id));
      await ctx.sync();
      items.forEach((b) => {
        if (!b.isNullObject) handlers.push(b.onDataChanged.add(onDataChanged));
      });
      await ctx.sync();
    });
  }

  async function onDataChanged() {
    clearTimeout(changeTimer);
    changeTimer = setTimeout(() => run(refresh), 300);
  }

  function saveSettings(now) {
    if (!inExcel) return;
    clearTimeout(saveTimer);
    const doSave = () => {
      const s = Office.context.document.settings;
      s.set(KEY_FMT, fmt);
      if (dataBindingId) s.set(KEY_DATA, dataBindingId);
      else s.remove(KEY_DATA);
      if (nodesBindingId) s.set(KEY_NODES, nodesBindingId);
      else s.remove(KEY_NODES);
      s.saveAsync(() => {});
    };
    if (now) doSave();
    else saveTimer = setTimeout(doSave, 600);
  }

  // ---------- Dane przykładowe ----------
  async function insertSample() {
    let tableName = null, sheetName = null;
    await Excel.run(async (ctx) => {
      let suffix = "";
      for (let i = 1; i < 50; i++) {
        suffix = i === 1 ? "" : " " + i;
        const ws = ctx.workbook.worksheets.getItemOrNullObject("Sankey dane" + suffix);
        const t = ctx.workbook.tables.getItemOrNullObject("SankeyDane" + suffix.trim());
        await ctx.sync();
        if (ws.isNullObject && t.isNullObject) break;
      }
      sheetName = "Sankey dane" + suffix;
      tableName = "SankeyDane" + suffix.trim();
      const sheet = ctx.workbook.worksheets.add(sheetName);
      const range = sheet.getRange("A1").getResizedRange(SAMPLE_DATA.length - 1, SAMPLE_DATA[0].length - 1);
      range.values = SAMPLE_DATA;
      const table = sheet.tables.add(range, true);
      table.name = tableName;
      table.style = "TableStyleLight9";
      sheet.getRange("C2:C" + SAMPLE_DATA.length).numberFormat = SAMPLE_DATA.slice(1).map(() => ["0.0"]);
      sheet.getRange("A:C").format.autofitColumns();
      await ctx.sync();

      if (dataBindingId) {
        const old = ctx.workbook.bindings.getItemOrNullObject(dataBindingId);
        await ctx.sync();
        if (!old.isNullObject) old.delete();
      }
      dataBindingId = newBindingId("data");
      await addBinding(ctx, null, tableName, dataBindingId);
    });
    fmt = Object.assign(clone(DEFAULT_FMT), clone(SAMPLE_FMT));
    selectedNode = null;
    saveSettings(true);
    $("empty").hidden = true;
    await refresh();
    await subscribe();
    showBanner("Dane przykładowe są w tabeli na arkuszu „" + sheetName + "”. Zmieniaj je, a wykres odświeży się sam.");
  }

  // ---------- Model ----------
  function str(x) {
    return x === null || x === undefined ? "" : String(x).trim();
  }

  function num(x) {
    if (x === "" || x === null || x === undefined || typeof x === "boolean") return null;
    const v = Number(x);
    return isNaN(v) ? null : v;
  }

  function norm(s) {
    return str(s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace("ł", "l");
  }

  // Nagłówek rozpoznawany po tekście w kolumnie wartości; kolumny po nazwach (PL/EN) albo po kolejności
  function readColumns(values, spec) {
    if (!values || !values.length) return { rows: [], header: false };
    const first = values[0];
    const known = [];
    Object.keys(spec.cols).forEach((k) => known.push.apply(known, spec.cols[k].names));
    const firstNorm = first.map(norm);
    const valueCell = spec.numericCol < first.length ? first[spec.numericCol] : null;
    const header =
      firstNorm.some((h) => known.indexOf(h) >= 0) ||
      (valueCell !== null && num(valueCell) === null && str(valueCell) !== "");
    const names = header ? firstNorm : [];
    const idx = {};
    Object.keys(spec.cols).forEach((key) => {
      const def = spec.cols[key];
      let i = header ? names.findIndex((h) => def.names.indexOf(h) >= 0) : -1;
      if (i < 0) i = def.pos < first.length ? def.pos : -1;
      idx[key] = i;
    });
    const rows = values.slice(header ? 1 : 0).map((r) => {
      const o = {};
      Object.keys(idx).forEach((k) => (o[k] = idx[k] >= 0 ? r[idx[k]] : ""));
      return o;
    });
    return { rows: rows, header: header };
  }

  const DATA_SPEC = {
    numericCol: 2,
    cols: {
      source: { pos: 0, names: ["source", "zrodlo", "from", "od", "z", "skad"] },
      target: { pos: 1, names: ["target", "cel", "to", "do", "dokad"] },
      value: { pos: 2, names: ["value", "wartosc", "amount", "kwota", "suma"] },
      color: { pos: 3, names: ["color", "kolor", "colour"] }
    }
  };
  const NODE_SPEC = {
    numericCol: 99,
    cols: {
      node: { pos: 0, names: ["node", "wezel", "nazwa", "name"] },
      color: { pos: 1, names: ["color", "kolor", "colour"] },
      note: { pos: 2, names: ["note", "dopisek", "opis", "notatka", "label"] },
      column: { pos: 3, names: ["column", "kolumna"] },
      display: { pos: 4, names: ["displayvalue", "display", "wartosc wyswietlana", "wyswietlana"] }
    }
  };

  function buildModel(dataValues, nodeValues, address, nodesAddress) {
    const parsed = readColumns(dataValues, DATA_SPEC);
    const order = [];
    const seen = new Set();
    const links = [];
    let skipped = 0;
    parsed.rows.forEach((r) => {
      const s = str(r.source), t = str(r.target), v = num(r.value);
      if (!s && !t && str(r.value) === "") return; // pusty wiersz
      if (!s || !t || v === null) {
        skipped++;
        return;
      }
      if (v <= 0) return;
      [s, t].forEach((n) => {
        if (!seen.has(n)) {
          seen.add(n);
          order.push(n);
        }
      });
      links.push({ source: s, target: t, value: v, color: str(r.color) });
    });

    const sheetNodes = {};
    if (nodeValues) {
      readColumns(nodeValues, NODE_SPEC).rows.forEach((r) => {
        const n = str(r.node);
        if (!n) return;
        sheetNodes[n] = { color: str(r.color), note: str(r.note), column: num(r.column), display: num(r.display) };
      });
    }

    const inSum = {}, outSum = {};
    links.forEach((l) => {
      outSum[l.source] = (outSum[l.source] || 0) + l.value;
      inSum[l.target] = (inSum[l.target] || 0) + l.value;
    });

    const nodes = order.map((name) => ({ name: name, value: Math.max(inSum[name] || 0, outSum[name] || 0), sheet: sheetNodes[name] || null }));
    return { graph: { nodes: nodes, links: links }, skipped: skipped, address: address, nodesAddress: nodesAddress || null };
  }

  // Właściwości węzła: arkusz węzłów > panel > paleta
  function nodeProps(node, index) {
    const pane = fmt.nodes[node.name] || {};
    const sh = node.sheet || {};
    const pal = PALETTES[fmt.palette] || PALETTES.office;
    const pick = (a, b) => (a !== null && a !== undefined && a !== "" ? a : b);
    return {
      color: pick(sh.color, pick(pane.color, pal[index % pal.length])),
      note: pick(sh.note, pane.note || ""),
      column: pick(sh.column, pick(pane.column, null)),
      display: pick(sh.display, pick(pane.display, null))
    };
  }

  // ---------- Rysowanie ----------
  function clone(o) {
    return JSON.parse(JSON.stringify(o));
  }

  function makeFormatter() {
    const locale = fmt.numberLocale === "en"
      ? d3.formatLocale({ decimal: ".", thousands: ",", grouping: [3], currency: ["", ""] })
      : d3.formatLocale({ decimal: ",", thousands: " ", grouping: [3], currency: ["", ""] });
    const f = locale.format(",." + Math.max(0, Math.min(6, Math.round(fmt.decimals))) + "f");
    return (v) => fmt.valuePrefix + f(v) + fmt.valueSuffix;
  }

  function labelLines(node, props, f) {
    const lines = [{ t: node.name, bold: true }];
    if (fmt.showValues) lines.push({ t: f(props.display !== null ? props.display : node.value), bold: false });
    if (props.note) lines.push({ t: props.note, bold: false });
    return lines;
  }

  function computeLayers(graph, columnOverride) {
    const aligners = { left: d3.sankeyLeft, right: d3.sankeyRight, center: d3.sankeyCenter, justify: d3.sankeyJustify };
    const base = aligners[fmt.nodeAlign] || d3.sankeyLeft;
    let probe;
    try {
      probe = d3.sankey().nodeId((d) => d.name)(cloneGraph(graph));
    } catch (e) {
      if (e && e.message === "circular link") throw new Error("Przepływy tworzą pętlę (np. A → B → A). Sankey wymaga układu bez cykli.");
      throw e;
    }
    const n = d3.max(probe.nodes, (d) => d.depth) + 1;
    const layers = new Map();
    probe.nodes.forEach((node) => {
      const o = columnOverride[node.name];
      const wanted = o !== null && o !== undefined ? o : base(node, n);
      layers.set(node.name, Math.max(0, Math.min(n - 1, Math.floor(wanted))));
    });
    const used = new Set(Array.from(layers.values()));
    const empty = [];
    for (let i = 0; i < n; i++) if (!used.has(i)) empty.push(i + 1);
    if (empty.length) throw new Error("Puste kolumny wykresu: " + empty.join(", ") + ". Popraw ustawienie „Kolumna” węzłów.");
    return { layers: layers, depth: n, probe: probe };
  }

  function cloneGraph(g) {
    return { nodes: g.nodes.map((d) => ({ name: d.name })), links: g.links.map((d) => Object.assign({}, d)) };
  }

  function draw() {
    if (!model) return;
    const box = $("chart");
    const W = Math.floor(box.clientWidth), H = Math.floor(box.clientHeight);
    if (W < 120 || H < 80) {
      clearChart();
      return;
    }
    const s = Math.min(1, W / 900, H / 500);
    try {
      lastSvg = drawSankey(box, W, H, W / s, H / s);
    } catch (e) {
      lastSvg = null;
      box.innerHTML = '<div class="chart-error"><b>Nie da się narysować wykresu</b><br>' + esc(e && e.message ? e.message : String(e)) + "</div>";
    }
  }

  function drawSankey(container, W, H, LW, LH) {
    const graph = model.graph;
    const f = makeFormatter();
    const props = {};
    graph.nodes.forEach((n, i) => (props[n.name] = nodeProps(n, i)));
    const columnOverride = {};
    Object.keys(props).forEach((k) => (columnOverride[k] = props[k].column));
    const lay = computeLayers(graph, columnOverride);
    model.depth = lay.depth;

    container.innerHTML = "";
    const svg = d3
      .select(container)
      .append("svg")
      .attr("width", W)
      .attr("height", H)
      .attr("viewBox", "0 0 " + LW + " " + LH)
      .attr("font-family", "Segoe UI, Roboto, Arial, sans-serif")
      .attr("font-size", fmt.fontSize);

    // Automatyczne marginesy: najdłuższa etykieta po lewej (źródła) i po prawej (ujścia)
    const lh = fmt.fontSize * 1.25;
    const meas = svg.append("text").attr("visibility", "hidden");
    const width = (t, bold) => {
      meas.attr("font-weight", bold ? 700 : 400).text(t);
      return meas.node().getComputedTextLength();
    };
    let maxLeft = 0, maxRight = 0, maxLines = 1;
    lay.probe.nodes.forEach((pn) => {
      const node = graph.nodes.find((x) => x.name === pn.name);
      const lines = labelLines(node, props[pn.name], f);
      maxLines = Math.max(maxLines, lines.length);
      const w = d3.max(lines, (l) => width(l.t, l.bold));
      const isSource = pn.targetLinks.length === 0, isSink = pn.sourceLinks.length === 0;
      if (isSource && !isSink && lay.layers.get(pn.name) === 0) maxLeft = Math.max(maxLeft, w);
      if (isSink) maxRight = Math.max(maxRight, w);
    });
    meas.remove();
    const titleH = fmt.showTitle && fmt.title ? fmt.titleSize * 1.8 : 0;
    const vpad = (maxLines * lh) / 2 + 6;
    const mLeft = Math.ceil(maxLeft) + 14, mRight = Math.ceil(maxRight) + 14;

    const sankey = d3
      .sankey()
      .nodeId((d) => d.name)
      .nodeAlign((node) => lay.layers.get(node.name))
      .nodeWidth(fmt.nodeWidth)
      .nodePadding(fmt.nodePadding)
      .iterations(32)
      .extent([[mLeft, titleH + vpad], [Math.max(mLeft + 50, LW - mRight), Math.max(titleH + vpad + 50, LH - vpad)]]);
    if (fmt.keepOrder) sankey.nodeSort(null);
    const laid = sankey(cloneGraph(graph));

    svg.insert("rect", ":first-child").attr("width", LW).attr("height", LH).attr("fill", fmt.background);

    if (titleH) {
      svg
        .append("text")
        .attr("x", LW / 2)
        .attr("y", titleH / 2 + 2)
        .attr("dy", "0.35em")
        .attr("text-anchor", "middle")
        .attr("font-size", fmt.titleSize)
        .attr("font-weight", 700)
        .attr("fill", fmt.titleColor)
        .text(fmt.title);
    }

    svg
      .append("g")
      .attr("fill", "none")
      .selectAll("path")
      .data(laid.links)
      .join("path")
      .attr("d", d3.sankeyLinkHorizontal())
      .attr("stroke", (d) => d.color || props[(fmt.linkColorBy === "source" ? d.source : d.target).name].color)
      .attr("stroke-opacity", fmt.linkOpacity)
      .attr("stroke-width", (d) => Math.max(1, d.width))
      .append("title")
      .text((d) => d.source.name + " → " + d.target.name + ": " + f(d.value));

    svg
      .append("g")
      .selectAll("rect")
      .data(laid.nodes)
      .join("rect")
      .attr("class", "node")
      .attr("x", (d) => d.x0)
      .attr("y", (d) => d.y0)
      .attr("width", (d) => d.x1 - d.x0)
      .attr("height", (d) => Math.max(1, d.y1 - d.y0))
      .attr("fill", (d) => props[d.name].color)
      .attr("stroke", (d) => (d.name === selectedNode ? "#222" : "none"))
      .attr("stroke-width", 2)
      .style("cursor", "pointer")
      .on("click", (ev, d) => {
        selectedNode = d.name;
        togglePane(true);
        draw();
      })
      .append("title")
      .text((d) => d.name + ": " + f(props[d.name].display !== null ? props[d.name].display : d.value) + "\nKliknij, aby sformatować węzeł");

    svg
      .append("g")
      .attr("pointer-events", "none")
      .selectAll("g")
      .data(laid.nodes)
      .join("g")
      .each(function (d) {
        const g = d3.select(this);
        const node = graph.nodes.find((x) => x.name === d.name);
        const lines = labelLines(node, props[d.name], f);
        const isSource = d.targetLinks.length === 0, isSink = d.sourceLinks.length === 0;
        const onLeft = isSink ? false : isSource ? true : d.x0 < LW / 2;
        const x = onLeft ? d.x0 - 6 : d.x1 + 6;
        const yTop = (d.y0 + d.y1) / 2 - ((lines.length - 1) * lh) / 2;
        const text = g.append("text").attr("text-anchor", onLeft ? "end" : "start").attr("fill", fmt.labelColor);
        lines.forEach((l, i) =>
          text
            .append("tspan")
            .attr("x", x)
            .attr("y", yTop + i * lh)
            .attr("dy", "0.35em")
            .attr("font-weight", l.bold ? 700 : 400)
            .text(l.t)
        );
        if (fmt.labelBackground && !isSource && !isSink) {
          const bb = text.node().getBBox();
          g.insert("rect", "text")
            .attr("x", bb.x - 4)
            .attr("y", bb.y - 2)
            .attr("width", bb.width + 8)
            .attr("height", bb.height + 4)
            .attr("rx", 3)
            .attr("fill", fmt.background)
            .attr("fill-opacity", 0.75);
        }
      });

    return svg.node();
  }

  function clearChart() {
    $("chart").innerHTML = "";
    lastSvg = null;
  }

  function scheduleDraw() {
    clearTimeout(drawTimer);
    drawTimer = setTimeout(() => run(async () => draw()), 60);
  }

  // ---------- Panel formatowania ----------
  const SECTIONS = [
    {
      id: "title", title: "Tytuł", fields: [
        { key: "showTitle", label: "Pokaż tytuł", type: "check" },
        { key: "title", label: "Tekst", type: "text" },
        { key: "titleSize", label: "Rozmiar", type: "number", min: 8, max: 60 },
        { key: "titleColor", label: "Kolor", type: "color" }
      ]
    },
    {
      id: "values", title: "Wartości", fields: [
        { key: "showValues", label: "Pokaż wartości w etykietach", type: "check" },
        { key: "valuePrefix", label: "Przed liczbą", type: "text", placeholder: "np. $" },
        { key: "valueSuffix", label: "Po liczbie", type: "text", placeholder: "np. mln zł" },
        { key: "decimals", label: "Miejsca po przecinku", type: "number", min: 0, max: 6 },
        { key: "numberLocale", label: "Format liczb", type: "select", options: [["pl", "1 234,5"], ["en", "1,234.5"]] }
      ]
    },
    {
      id: "nodesStyle", title: "Węzły", fields: [
        { key: "palette", label: "Paleta kolorów", type: "select", options: [["office", "Office"], ["office2013", "Office 2013-2022"], ["pastel", "Pastelowa"], ["mono", "Szara"]] },
        { key: "nodeWidth", label: "Szerokość", type: "range", min: 4, max: 60, step: 1 },
        { key: "nodePadding", label: "Odstęp między węzłami", type: "range", min: 0, max: 80, step: 1 },
        { key: "nodeAlign", label: "Wyrównanie kolumn", type: "select", options: [["left", "Do lewej"], ["right", "Do prawej"], ["justify", "Wyjustuj"], ["center", "Do środka"]] },
        { key: "keepOrder", label: "Kolejność jak w danych", type: "check" }
      ]
    },
    {
      id: "links", title: "Przepływy", fields: [
        { key: "linkColorBy", label: "Kolor przepływu", type: "select", options: [["target", "jak węzeł docelowy"], ["source", "jak węzeł źródłowy"]] },
        { key: "linkOpacity", label: "Krycie", type: "range", min: 0.1, max: 1, step: 0.05 }
      ]
    },
    {
      id: "labels", title: "Etykiety i tło", fields: [
        { key: "fontSize", label: "Rozmiar czcionki", type: "number", min: 7, max: 32 },
        { key: "labelColor", label: "Kolor etykiet", type: "color" },
        { key: "labelBackground", label: "Tło pod etykietami w środku", type: "check" },
        { key: "background", label: "Tło wykresu", type: "color" }
      ]
    }
  ];

  function renderPane() {
    const openState = {};
    $("pane-body").querySelectorAll("details").forEach((d) => (openState[d.dataset.id] = d.open));
    const isOpen = (id, def) => (openState[id] !== undefined ? openState[id] : def);
    let html = "";

    // Dane
    const dis = inExcel ? "" : " disabled";
    const addr = model ? model.address : "(brak)";
    html +=
      '<details data-id="data"' + (isOpen("data", !model) ? " open" : "") + "><summary>Dane</summary>" +
      '<div class="field"><span class="lbl">Zakres</span><code>' + esc(addr) + "</code></div>" +
      '<div class="btns"><button data-act="use-data" class="primary"' + dis + '>Użyj zaznaczenia</button><button data-act="sample"' + dis + '>Dane przykładowe</button></div>' +
      '<p class="hint">Kolumny: Źródło, Cel, Wartość i opcjonalnie Kolor (#RRGGBB) przepływu. Komórka w tabeli Excela = cała tabela, nowe wiersze dochodzą same.</p>' +
      "</details>";

    SECTIONS.forEach((sec) => {
      html += '<details data-id="' + sec.id + '"' + (isOpen(sec.id, sec.id === "title") ? " open" : "") + "><summary>" + sec.title + "</summary>";
      sec.fields.forEach((fd) => (html += fieldHtml(fd)));
      html += "</details>";
    });

    // Pojedyncze węzły
    html += '<details data-id="nodes"' + (isOpen("nodes", !!selectedNode) || selectedNode ? " open" : "") + "><summary>Pojedyncze węzły</summary>";
    if (model) {
      html += '<p class="hint">Kliknij węzeł na wykresie, żeby od razu tu trafić. Kolumna liczona od 1.</p>';
      model.graph.nodes.forEach((n, i) => {
        const p = nodeProps(n, i);
        const pane = fmt.nodes[n.name] || {};
        const fromSheet = n.sheet ? ' <span class="tag">z arkusza</span>' : "";
        const cols = ['<option value="">auto</option>'];
        for (let c = 0; c < (model.depth || 1); c++) cols.push('<option value="' + c + '"' + (pane.column === c ? " selected" : "") + ">" + (c + 1) + "</option>");
        html +=
          '<div class="node-row' + (n.name === selectedNode ? " selected" : "") + '" data-node="' + esc(n.name) + '">' +
          '<div class="node-head"><input type="color" data-nkey="color" value="' + toHex(p.color) + '"><b>' + esc(n.name) + "</b>" + fromSheet + "</div>" +
          '<div class="swatches">' + QUICK_COLORS.map((c) => '<button class="sw" data-act="swatch" data-color="' + c + '" style="background:' + c + '" title="' + c + '"></button>').join("") +
          '<button class="sw reset" data-act="reset-color" title="Kolor z palety">↺</button></div>' +
          '<div class="grid2">' +
          '<label>Dopisek<input type="text" data-nkey="note" value="' + esc(pane.note || "") + '" placeholder="np. 7% r/r"></label>' +
          '<label>Kolumna<select data-nkey="column">' + cols.join("") + "</select></label>" +
          '<label>Wartość wyświetlana<input type="number" step="any" data-nkey="display" value="' + (pane.display !== null && pane.display !== undefined ? pane.display : "") + '" placeholder="auto"></label>' +
          "</div></div>";
      });
      html += '<div class="btns"><button data-act="reset-nodes">Wyczyść ustawienia węzłów</button></div>';
    }
    html += "</details>";

    // Zaawansowane
    html +=
      '<details data-id="adv"' + (isOpen("adv", false) ? " open" : "") + "><summary>Zaawansowane</summary>" +
      '<p class="hint">Opcjonalny zakres z opisami węzłów liczonymi formułami: Węzeł | Kolor | Dopisek | Kolumna | Wartość wyświetlana. Ma pierwszeństwo przed ustawieniami powyżej.</p>' +
      '<div class="field"><span class="lbl">Zakres węzłów</span><code>' + esc(model && model.nodesAddress ? model.nodesAddress : "(brak)") + "</code></div>" +
      '<div class="btns"><button data-act="use-nodes"' + dis + '>Użyj zaznaczenia</button>' + (nodesBindingId ? '<button data-act="unbind-nodes">Odłącz</button>' : "") + "</div>" +
      '<p class="hint">Wersja dodatku: ' + VERSION + "</p>" +
      "</details>";

    $("pane-body").innerHTML = html;
    if (selectedNode) {
      const row = $("pane-body").querySelector('.node-row.selected');
      if (row) row.scrollIntoView({ block: "center" });
    }
  }

  function fieldHtml(fd) {
    const v = fmt[fd.key];
    const id = "f-" + fd.key;
    if (fd.type === "check")
      return '<label class="check"><input type="checkbox" id="' + id + '" data-key="' + fd.key + '"' + (v ? " checked" : "") + "> " + fd.label + "</label>";
    let input;
    if (fd.type === "select")
      input = '<select id="' + id + '" data-key="' + fd.key + '">' + fd.options.map((o) => '<option value="' + o[0] + '"' + (String(v) === o[0] ? " selected" : "") + ">" + o[1] + "</option>").join("") + "</select>";
    else if (fd.type === "range")
      input = '<span class="range"><input type="range" id="' + id + '" data-key="' + fd.key + '" min="' + fd.min + '" max="' + fd.max + '" step="' + fd.step + '" value="' + v + '"><output>' + v + "</output></span>";
    else if (fd.type === "color") input = '<input type="color" id="' + id + '" data-key="' + fd.key + '" value="' + toHex(v) + '">';
    else if (fd.type === "number") input = '<input type="number" id="' + id + '" data-key="' + fd.key + '" min="' + fd.min + '" max="' + fd.max + '" value="' + v + '">';
    else input = '<input type="text" id="' + id + '" data-key="' + fd.key + '" value="' + esc(v) + '" placeholder="' + esc(fd.placeholder || "") + '">';
    return '<div class="field"><label class="lbl" for="' + id + '">' + fd.label + "</label>" + input + "</div>";
  }

  function onPaneInput(ev) {
    const el = ev.target;
    if (el.dataset.key) {
      const key = el.dataset.key;
      const def = DEFAULT_FMT[key];
      let v;
      if (el.type === "checkbox") v = el.checked;
      else if (typeof def === "number") {
        v = Number(el.value);
        if (isNaN(v)) return;
      } else v = el.value;
      fmt[key] = v;
      if (el.type === "range" && el.nextElementSibling) el.nextElementSibling.textContent = el.value;
      if (key === "title" && ev.type === "input" && v && !fmt.showTitle) {
        fmt.showTitle = true;
        const c = $("f-showTitle");
        if (c) c.checked = true;
      }
      saveSettings(false);
      scheduleDraw();
      if (key === "palette" && ev.type === "change") renderPane();
      return;
    }
    if (el.dataset.nkey) {
      const row = el.closest(".node-row");
      if (!row) return;
      const name = row.dataset.node;
      const k = el.dataset.nkey;
      const entry = fmt.nodes[name] || (fmt.nodes[name] = {});
      if (k === "color") entry.color = el.value;
      if (k === "note") entry.note = el.value;
      if (k === "column") entry.column = el.value === "" ? null : Number(el.value);
      if (k === "display") entry.display = el.value === "" ? null : Number(el.value);
      selectedNode = name;
      saveSettings(false);
      scheduleDraw();
    }
  }

  function onPaneClick(ev) {
    const btn = ev.target.closest("button");
    if (!btn) return;
    const act = btn.dataset.act;
    const row = btn.closest(".node-row");
    if (act === "use-data") run(useSelectionClicked, true);
    if (act === "sample") run(insertSample, true);
    if (act === "use-nodes")
      run(async () => {
        const ok = await bindFromSelection("nodes", false);
        if (ok) {
          await refresh();
          await subscribe();
          renderPane();
        }
      }, true);
    if (act === "unbind-nodes") run(async () => { await unbindNodes(); renderPane(); }, true);
    if ((act === "swatch" || act === "reset-color") && row) {
      const name = row.dataset.node;
      const entry = fmt.nodes[name] || (fmt.nodes[name] = {});
      if (act === "swatch") entry.color = btn.dataset.color;
      else delete entry.color;
      selectedNode = name;
      saveSettings(false);
      draw();
      renderPane();
    }
    if (act === "reset-nodes") {
      fmt.nodes = {};
      selectedNode = null;
      saveSettings(false);
      draw();
      renderPane();
    }
  }

  // ---------- Kopia jako obraz ----------
  function svgToPngBase64(svgEl, w, h, scale) {
    return new Promise((resolve, reject) => {
      const clone = svgEl.cloneNode(true);
      clone.querySelectorAll("rect.node").forEach((r) => r.setAttribute("stroke", "none"));
      const xml = new XMLSerializer().serializeToString(clone);
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(w * scale);
        canvas.height = Math.round(h * scale);
        const g = canvas.getContext("2d");
        g.scale(scale, scale);
        g.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/png").split(",")[1]);
      };
      img.onerror = () => reject(new Error("Nie udało się zamienić wykresu na obraz."));
      img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(xml);
    });
  }

  async function insertPngCopy() {
    const keep = selectedNode;
    selectedNode = null;
    draw();
    selectedNode = keep;
    if (!lastSvg) throw new Error("Brak wykresu do skopiowania.");
    const w = Number(lastSvg.getAttribute("width")), h = Number(lastSvg.getAttribute("height"));
    const png = await svgToPngBase64(lastSvg, w, h, fmt.imageScale || 2);
    await Excel.run(async (ctx) => {
      const cell = ctx.workbook.getActiveCell().load("left,top");
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      await ctx.sync();
      const shape = ws.shapes.addImage(png);
      shape.name = "Sankey (obraz)";
      shape.left = cell.left;
      shape.top = cell.top;
      shape.lockAspectRatio = false;
      shape.width = w * 0.75;
      shape.height = h * 0.75;
      shape.lockAspectRatio = true;
      await ctx.sync();
    });
    showBanner("Wstawiono kopię jako obraz przy aktywnej komórce. Możesz ją skopiować np. do PowerPointa.");
  }

  // ---------- Pomocnicze ----------
  function toHex(c) {
    const s = str(c);
    if (/^#[0-9a-f]{6}$/i.test(s)) return s;
    if (/^#[0-9a-f]{3}$/i.test(s)) return "#" + s.slice(1).split("").map((x) => x + x).join("");
    const d = d3.color(s);
    return d ? d.formatHex() : "#6b6b6b";
  }

  function showBanner(text, isError) {
    const b = $("banner");
    b.textContent = text;
    b.className = isError ? "error" : "";
    b.hidden = false;
    clearTimeout(b._t);
    b._t = setTimeout(() => (b.hidden = true), isError ? 9000 : 6000);
  }

  function esc(s) {
    return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
  }

  // soft = true: błąd w pasku, wykres zostaje
  async function run(fn, soft) {
    try {
      await fn();
    } catch (e) {
      console.error(e);
      const loc = e && e.debugInfo && e.debugInfo.errorLocation ? " [" + e.debugInfo.errorLocation + "]" : "";
      const msg = (e && e.message ? e.message : String(e)) + loc;
      showBanner("Błąd: " + msg, true);
      if (!soft && !lastSvg && !model) showEmpty();
    }
  }
})();
