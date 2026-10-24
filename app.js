/* Sankey Chart: content add-in dla Excela
   Dane: tabela przepływów (Source | Target | Value | Color), opcjonalnie tabela węzłów
   (Node | Color | Note | Column | DisplayValue) i tabela parametrów (Parameter | Value | Opis).
   Każdy wstawiony wykres pamięta własne nazwy tabel (ustawienia dodatku w pliku). */
(function () {
  "use strict";

  const SETTINGS_KEY = "sankeyTables";
  const DEFAULT_TABLES = { links: "SankeyLinks", nodes: "SankeyNodes", config: "SankeyConfig" };

  // Parametry nadpisywane w tabeli parametrów. Typ wartości domyślnej = typ parametru.
  const DEFAULTS = {
    title: "",
    marginTop: 20,
    marginRight: 130,
    marginBottom: 30,
    marginLeft: 130,
    nodeWidth: 24,
    nodePadding: 36,
    nodeAlign: "left",
    keepOrder: true,
    iterations: 32,
    linkOpacity: 0.45,
    defaultNodeColor: "#6b6b6b",
    background: "#ffffff",
    fontFamily: "Segoe UI, Roboto, Arial, sans-serif",
    fontSize: 12,
    labelColor: "#222222",
    labelBackground: true,
    titleSize: 24,
    titleColor: "#1f4e79",
    valuePrefix: "",
    valueSuffix: "",
    decimals: 1,
    numberLocale: "en",
    imageScale: 2,
    minWidth: 900,
    minHeight: 500
  };

  // ---------- Dane przykładowe (Apple FY22) ----------
  const LINKS_SAMPLE = [
    ["Source", "Target", "Value", "Color"],
    ["iPhone", "Products", 205.5, ""],
    ["Mac", "Products", 40.2, ""],
    ["iPad", "Products", 29.3, ""],
    ["Wearables", "Products", 41.2, ""],
    ["Products", "Revenue", 316.2, ""],
    ["Services", "Revenue", 78.2, ""],
    ["Revenue", "Gross profit", 170.9, ""],
    ["Revenue", "Cost of revenue", 223.5, ""],
    ["Gross profit", "Operating profit", 119.5, ""],
    ["Gross profit", "Operating expenses", 51.4, ""],
    ["Operating profit", "Net profit", 99.8, ""],
    ["Operating profit", "Tax", 19.3, ""],
    ["Operating profit", "Other", 0.3, ""],
    ["Operating expenses", "R&D", 26.3, ""],
    ["Operating expenses", "SG&A", 25.1, ""],
    ["Cost of revenue", "Product costs", 201.4, ""],
    ["Cost of revenue", "Service costs", 22.1, ""]
  ];

  const GRAY = "#6b6b6b", GREEN = "#00a800", RED = "#b30000";
  const NODES_SAMPLE = [
    ["Node", "Color", "Note", "Column", "DisplayValue"],
    ["iPhone", GRAY, "7% Y/Y", "", ""],
    ["Mac", GRAY, "14% Y/Y", "", ""],
    ["iPad", GRAY, "-8% Y/Y", "", ""],
    ["Wearables", GRAY, "7% Y/Y", "", ""],
    ["Products", GRAY, "6% Y/Y", "", ""],
    ["Services", GRAY, "14% Y/Y", 1, ""],
    ["Revenue", GRAY, "8% Y/Y", "", 394.3],
    ["Gross profit", GREEN, "12% Y/Y", "", ""],
    ["Cost of revenue", RED, "5% Y/Y", "", ""],
    ["Operating profit", GREEN, "10% Y/Y", "", ""],
    ["Operating expenses", RED, "17% Y/Y", "", ""],
    ["Product costs", RED, "5% Y/Y", "", ""],
    ["Service costs", RED, "7% Y/Y", "", ""],
    ["Net profit", GREEN, "5% Y/Y", "", ""],
    ["Tax", RED, "33% Y/Y", "", ""],
    ["Other", RED, "", "", ""],
    ["R&D", RED, "20% Y/Y", "", ""],
    ["SG&A", RED, "14% Y/Y", "", ""]
  ];

  const CONFIG_SAMPLE = [
    ["Parameter", "Value", "Opis"],
    ["title", "Apple FY22 Income Statement", "Tytuł wykresu (puste = brak tytułu)"],
    ["marginTop", 70, "Margines górny w px (miejsce na tytuł; bez tytułu ok. 20)"],
    ["marginRight", 130, "Margines prawy (etykiety ostatniej kolumny)"],
    ["marginBottom", 30, "Margines dolny"],
    ["marginLeft", 130, "Margines lewy (etykiety pierwszej kolumny)"],
    ["nodeWidth", 24, "Szerokość prostokąta węzła"],
    ["nodePadding", 36, "Minimalny odstęp pionowy między węzłami"],
    ["nodeAlign", "left", "left | right | center | justify (kolumna Column w tabeli węzłów ma pierwszeństwo)"],
    ["keepOrder", true, "TRUE = kolejność węzłów w kolumnie jak w tabeli węzłów, FALSE = automatyczna"],
    ["iterations", 32, "Liczba iteracji układu (więcej = mniej krzyżowań)"],
    ["linkOpacity", 0.45, "Przezroczystość przepływów 0-1"],
    ["defaultNodeColor", "#6b6b6b", "Kolor węzła bez koloru w tabeli węzłów"],
    ["background", "#ffffff", "Kolor tła"],
    ["fontFamily", "Segoe UI, Roboto, Arial, sans-serif", "Czcionka"],
    ["fontSize", 12, "Rozmiar czcionki etykiet"],
    ["labelColor", "#222222", "Kolor etykiet"],
    ["labelBackground", true, "Półprzezroczyste tło pod etykietami węzłów środkowych"],
    ["titleSize", 24, "Rozmiar tytułu"],
    ["titleColor", "#1f4e79", "Kolor tytułu"],
    ["valuePrefix", "$", "Tekst przed wartością"],
    ["valueSuffix", "B", "Tekst po wartości"],
    ["decimals", 1, "Liczba miejsc po przecinku"],
    ["numberLocale", "en", "en = 1,234.5 | pl = 1 234,5"],
    ["imageScale", 2, "Rozdzielczość kopii PNG (2 = ostra na ekranach HiDPI)"],
    ["minWidth", 900, "Poniżej tej szerokości obiektu wykres jest skalowany proporcjonalnie (razem z tekstem)"],
    ["minHeight", 500, "Poniżej tej wysokości obiektu wykres jest skalowany proporcjonalnie"]
  ];

  // ---------- Stan ----------
  let inExcel = false;
  let tables = Object.assign({}, DEFAULT_TABLES);
  let handlers = [];
  let cache = null; // { graph, cfg, unknown }
  let lastSvg = null;
  let changeTimer = null;
  let resizeTimer = null;

  const $ = (id) => document.getElementById(id);

  Office.onReady((info) => {
    inExcel = info && info.host === Office.HostType.Excel;
    wireUi();
    if (!inExcel) {
      // Otwarte w zwykłej przeglądarce: pokaż demo, żeby sprawdzić, że hosting działa.
      $("btn-copy").hidden = true;
      $("btn-settings").hidden = true;
      cache = prepare(rowsToObjects(LINKS_SAMPLE), rowsToObjects(NODES_SAMPLE), rowsToObjects(CONFIG_SAMPLE));
      draw();
      showBanner("Strona dodatku działa. To jest podgląd demo; w Excelu wykres czyta dane z Twoich tabel.");
      return;
    }
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) $("btn-copy").hidden = true;
    const saved = Office.context.document.settings.get(SETTINGS_KEY);
    if (saved && saved.links) tables = Object.assign({}, DEFAULT_TABLES, saved);
    run(async () => {
      await refresh();
      await subscribe();
    });
  });

  function wireUi() {
    $("btn-refresh").addEventListener("click", () => run(inExcel ? refresh : async () => draw()));
    $("btn-settings").addEventListener("click", () => run(openSettings, true));
    $("btn-copy").addEventListener("click", () => run(insertPngCopy, true));
    $("btn-save").addEventListener("click", () => run(saveSettings));
    $("btn-cancel").addEventListener("click", () => ($("settings").hidden = true));
    $("btn-sample").addEventListener("click", () => run(createSample));
    new ResizeObserver(() => {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(() => run(async () => draw()), 120);
    }).observe($("chart"));
  }

  // ---------- Odczyt danych z Excela ----------
  async function readTable(ctx, name) {
    if (!name) return null;
    const table = ctx.workbook.tables.getItemOrNullObject(name);
    await ctx.sync();
    if (table.isNullObject) return null;
    const header = table.getHeaderRowRange().load("values");
    const body = table.getDataBodyRange().load("values");
    await ctx.sync();
    return rowsToObjects([header.values[0]].concat(body.values));
  }

  function rowsToObjects(rows) {
    const cols = rows[0].map((h) => String(h).trim().toLowerCase());
    return rows.slice(1).map((row) => {
      const o = {};
      cols.forEach((c, i) => (o[c] = row[i]));
      return o;
    });
  }

  async function refresh() {
    await Excel.run(async (ctx) => {
      const linkRows = await readTable(ctx, tables.links);
      if (!linkRows) {
        cache = null;
        clearChart();
        showMessage(
          "Nie znaleziono tabeli przepływów <b>" + esc(tables.links) + "</b>.",
          "Utwórz dane przykładowe albo wskaż własne tabele w ustawieniach.",
          true
        );
        return;
      }
      const nodeRows = (await readTable(ctx, tables.nodes)) || [];
      const confRows = (await readTable(ctx, tables.config)) || [];
      cache = prepare(linkRows, nodeRows, confRows);
    });
    if (cache) draw();
  }

  function prepare(linkRows, nodeRows, confRows) {
    const parsed = parseConfig(confRows);
    const graph = buildGraph(linkRows, nodeRows);
    if (!graph.links.length) throw new Error("Tabela przepływów nie zawiera wierszy z wartością > 0.");
    return { graph: graph, cfg: parsed.cfg, unknown: parsed.unknown };
  }

  // Nasłuch zmian w tabelach: każda edycja odświeża wykres
  async function subscribe() {
    for (const h of handlers) {
      try {
        await Excel.run(h.context, async (ctx) => {
          h.remove();
          await ctx.sync();
        });
      } catch (e) {
        /* tabela mogła zniknąć */
      }
    }
    handlers = [];
    await Excel.run(async (ctx) => {
      const names = [tables.links, tables.nodes, tables.config].filter(Boolean);
      const items = names.map((n) => ctx.workbook.tables.getItemOrNullObject(n));
      await ctx.sync();
      items.forEach((t) => {
        if (!t.isNullObject) handlers.push(t.onChanged.add(onTableChanged));
      });
      await ctx.sync();
    });
  }

  async function onTableChanged() {
    clearTimeout(changeTimer);
    changeTimer = setTimeout(() => run(refresh), 400);
  }

  // ---------- Parsowanie ----------
  function str(x) {
    return x === null || x === undefined ? "" : String(x).trim();
  }

  function num(x) {
    if (x === "" || x === null || x === undefined || typeof x === "boolean") return null;
    const v = Number(x);
    return isNaN(v) ? null : v;
  }

  function parseConfig(rows) {
    const cfg = Object.assign({}, DEFAULTS);
    const unknown = [];
    const keys = Object.keys(DEFAULTS);
    (rows || []).forEach((r) => {
      const key = str(r.parameter);
      if (!key) return;
      const match = keys.filter((k) => k.toLowerCase() === key.toLowerCase())[0];
      if (!match) {
        unknown.push(key);
        return;
      }
      const def = DEFAULTS[match];
      let v = r.value;
      if (typeof def === "number") {
        v = num(v);
        if (v === null) return;
      } else if (typeof def === "boolean") {
        v = v === true || v === 1 || ["TRUE", "PRAWDA", "1", "TAK", "YES"].indexOf(str(v).toUpperCase()) >= 0;
      } else {
        v = str(v);
      }
      cfg[match] = v;
    });
    return { cfg: cfg, unknown: unknown };
  }

  function buildGraph(linkRows, nodeRows) {
    const meta = new Map();
    const order = [];
    (nodeRows || []).forEach((r) => {
      const name = str(r.node);
      if (!name || meta.has(name)) return;
      meta.set(name, { color: str(r.color), note: str(r.note), column: num(r.column), display: num(r.displayvalue) });
      order.push(name);
    });
    const links = [];
    const used = new Set();
    linkRows.forEach((r) => {
      const s = str(r.source), t = str(r.target), v = num(r.value);
      if (!s || !t || v === null || v <= 0) return;
      links.push({ source: s, target: t, value: v, color: str(r.color) });
      used.add(s);
      used.add(t);
    });
    links.forEach((l) =>
      [l.source, l.target].forEach((n) => {
        if (!meta.has(n)) {
          meta.set(n, { color: "", note: "", column: null, display: null });
          order.push(n);
        }
      })
    );
    const nodes = order.filter((n) => used.has(n)).map((n) => Object.assign({ name: n }, meta.get(n)));
    return { nodes: nodes, links: links };
  }

  // ---------- Rysowanie ----------
  function cloneGraph(g) {
    return {
      nodes: g.nodes.map((d) => Object.assign({}, d)),
      links: g.links.map((d) => Object.assign({}, d))
    };
  }

  function computeLayers(graph, cfg) {
    const aligners = { left: d3.sankeyLeft, right: d3.sankeyRight, center: d3.sankeyCenter, justify: d3.sankeyJustify };
    const base = aligners[str(cfg.nodeAlign).toLowerCase()] || d3.sankeyLeft;
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
      const wanted = node.column !== null && node.column !== undefined ? node.column : base(node, n);
      layers.set(node.name, Math.max(0, Math.min(n - 1, Math.floor(wanted))));
    });
    const usedCols = new Set(Array.from(layers.values()));
    const empty = [];
    for (let i = 0; i < n; i++) if (!usedCols.has(i)) empty.push(i);
    if (empty.length) {
      throw new Error("Puste kolumny wykresu: " + empty.join(", ") + ". Popraw wartości w kolumnie Column tabeli węzłów (numeracja od 0).");
    }
    return layers;
  }

  function draw() {
    if (!cache) return;
    const box = $("chart");
    const W = Math.floor(box.clientWidth), H = Math.floor(box.clientHeight);
    if (W < 120 || H < 80) {
      clearChart();
      showMessage("Obiekt jest za mały, żeby narysować wykres.", "Powiększ go, przeciągając za róg.", false);
      return;
    }
    hideMessage();
    // Mały obiekt: układ liczony w większej przestrzeni i zmniejszany proporcjonalnie, żeby etykiety się nie zlewały
    const s = Math.min(1, W / Math.max(1, cache.cfg.minWidth), H / Math.max(1, cache.cfg.minHeight));
    lastSvg = drawSankey(box, cache.graph, cache.cfg, W, H, W / s, H / s);
    if (cache.unknown.length) showBanner("Nieznane parametry: " + cache.unknown.join(", "));
  }

  // W, H: rozmiar obiektu w px; LW, LH: przestrzeń układu (viewBox)
  function drawSankey(container, graph, cfg, W, H, LW, LH) {
    const layers = computeLayers(graph, cfg);
    const sankey = d3
      .sankey()
      .nodeId((d) => d.name)
      .nodeAlign((node) => layers.get(node.name))
      .nodeWidth(cfg.nodeWidth)
      .nodePadding(cfg.nodePadding)
      .iterations(cfg.iterations)
      .extent([[cfg.marginLeft, cfg.marginTop], [LW - cfg.marginRight, LH - cfg.marginBottom]]);
    if (cfg.keepOrder) sankey.nodeSort(null);
    const laid = sankey(cloneGraph(graph));
    const nodes = laid.nodes, links = laid.links;

    const locale = str(cfg.numberLocale).toLowerCase() === "pl"
      ? d3.formatLocale({ decimal: ",", thousands: " ", grouping: [3], currency: ["", ""] })
      : d3.formatLocale({ decimal: ".", thousands: ",", grouping: [3], currency: ["", ""] });
    const fmtNum = locale.format(",." + Math.max(0, Math.round(cfg.decimals)) + "f");
    const fmt = (v) => cfg.valuePrefix + fmtNum(v) + cfg.valueSuffix;

    container.innerHTML = "";
    const svg = d3
      .select(container)
      .append("svg")
      .attr("width", W)
      .attr("height", H)
      .attr("viewBox", "0 0 " + LW + " " + LH)
      .attr("font-family", cfg.fontFamily)
      .attr("font-size", cfg.fontSize);

    svg.append("rect").attr("width", LW).attr("height", LH).attr("fill", cfg.background);

    if (cfg.title) {
      svg
        .append("text")
        .attr("x", LW / 2)
        .attr("y", cfg.marginTop / 2)
        .attr("dy", "0.35em")
        .attr("text-anchor", "middle")
        .attr("font-size", cfg.titleSize)
        .attr("font-weight", 700)
        .attr("fill", cfg.titleColor)
        .text(cfg.title);
    }

    // Przepływy: kolor z kolumny Color, a gdy pusty, kolor węzła docelowego
    svg
      .append("g")
      .attr("fill", "none")
      .selectAll("path")
      .data(links)
      .join("path")
      .attr("d", d3.sankeyLinkHorizontal())
      .attr("stroke", (d) => d.color || d.target.color || cfg.defaultNodeColor)
      .attr("stroke-opacity", cfg.linkOpacity)
      .attr("stroke-width", (d) => Math.max(1, d.width))
      .append("title")
      .text((d) => d.source.name + " → " + d.target.name + ": " + fmt(d.value));

    svg
      .append("g")
      .selectAll("rect")
      .data(nodes)
      .join("rect")
      .attr("x", (d) => d.x0)
      .attr("y", (d) => d.y0)
      .attr("width", (d) => d.x1 - d.x0)
      .attr("height", (d) => Math.max(1, d.y1 - d.y0))
      .attr("fill", (d) => d.color || cfg.defaultNodeColor)
      .append("title")
      .text((d) => d.name + ": " + fmt(d.display !== null && d.display !== undefined ? d.display : d.value));

    // Etykiety: źródła po lewej, ujścia po prawej, węzły środkowe zależnie od połowy wykresu
    const lh = cfg.fontSize * 1.25;
    svg
      .append("g")
      .selectAll("g")
      .data(nodes)
      .join("g")
      .each(function (d) {
        const g = d3.select(this);
        const isSource = d.targetLinks.length === 0;
        const isSink = d.sourceLinks.length === 0;
        const onLeft = isSink ? false : isSource ? true : d.x0 < LW / 2;
        const x = onLeft ? d.x0 - 6 : d.x1 + 6;
        const shown = d.display !== null && d.display !== undefined ? d.display : d.value;
        const lines = [{ t: d.name, bold: true }, { t: fmt(shown), bold: false }];
        if (d.note) lines.push({ t: d.note, bold: false });
        const yTop = (d.y0 + d.y1) / 2 - ((lines.length - 1) * lh) / 2;
        const text = g.append("text").attr("text-anchor", onLeft ? "end" : "start").attr("fill", cfg.labelColor);
        lines.forEach((l, i) =>
          text
            .append("tspan")
            .attr("x", x)
            .attr("y", yTop + i * lh)
            .attr("dy", "0.35em")
            .attr("font-weight", l.bold ? 700 : 400)
            .text(l.t)
        );
        if (cfg.labelBackground && !isSource && !isSink) {
          const bb = text.node().getBBox();
          g.insert("rect", "text")
            .attr("x", bb.x - 4)
            .attr("y", bb.y - 2)
            .attr("width", bb.width + 8)
            .attr("height", bb.height + 4)
            .attr("rx", 3)
            .attr("fill", cfg.background)
            .attr("fill-opacity", 0.75);
        }
      });

    return svg.node();
  }

  function svgToPngBase64(svgEl, w, h, scale) {
    return new Promise((resolve, reject) => {
      const xml = new XMLSerializer().serializeToString(svgEl);
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

  // ---------- Akcje ----------
  async function insertPngCopy() {
    if (!lastSvg || !cache) throw new Error("Brak wykresu do skopiowania.");
    const w = Number(lastSvg.getAttribute("width")), h = Number(lastSvg.getAttribute("height"));
    const png = await svgToPngBase64(lastSvg, w, h, cache.cfg.imageScale);
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

  async function openSettings() {
    let names = [];
    await Excel.run(async (ctx) => {
      const col = ctx.workbook.tables.load("items/name");
      await ctx.sync();
      names = col.items.map((t) => t.name).sort();
    });
    fillSelect($("sel-links"), names, tables.links, false);
    fillSelect($("sel-nodes"), names, tables.nodes, true);
    fillSelect($("sel-config"), names, tables.config, true);
    $("settings").hidden = false;
  }

  function fillSelect(sel, names, current, optional) {
    sel.innerHTML = "";
    if (optional) sel.appendChild(new Option("(brak)", ""));
    const list = names.slice();
    if (current && list.indexOf(current) < 0) list.unshift(current);
    list.forEach((n) => sel.appendChild(new Option(names.indexOf(n) < 0 ? n + " (nie istnieje)" : n, n)));
    sel.value = current || "";
  }

  async function saveSettings() {
    const links = $("sel-links").value;
    if (!links) throw new Error("Wybierz tabelę przepływów.");
    await applyTables({ links: links, nodes: $("sel-nodes").value, config: $("sel-config").value });
    $("settings").hidden = true;
  }

  async function applyTables(next) {
    tables = next;
    Office.context.document.settings.set(SETTINGS_KEY, tables);
    await new Promise((resolve) => Office.context.document.settings.saveAsync(() => resolve()));
    await refresh();
    await subscribe();
  }

  // Tworzy arkusz z trzema tabelami przykładowymi o wolnych nazwach i podpina je do tego wykresu
  async function createSample() {
    let created = null;
    await Excel.run(async (ctx) => {
      let suffix = "";
      for (let i = 1; i < 50; i++) {
        suffix = i === 1 ? "" : String(i);
        const probes = [
          ctx.workbook.tables.getItemOrNullObject("SankeyLinks" + suffix),
          ctx.workbook.tables.getItemOrNullObject("SankeyNodes" + suffix),
          ctx.workbook.tables.getItemOrNullObject("SankeyConfig" + suffix),
          ctx.workbook.worksheets.getItemOrNullObject("Sankey dane" + (suffix ? " " + suffix : ""))
        ];
        await ctx.sync();
        if (probes.every((p) => p.isNullObject)) break;
      }
      const sheetName = "Sankey dane" + (suffix ? " " + suffix : "");
      const sheet = ctx.workbook.worksheets.add(sheetName);
      created = { links: "SankeyLinks" + suffix, nodes: "SankeyNodes" + suffix, config: "SankeyConfig" + suffix, sheet: sheetName };
      addTable(sheet, "A1", created.links, LINKS_SAMPLE);
      addTable(sheet, "F1", created.nodes, NODES_SAMPLE);
      addTable(sheet, "L1", created.config, CONFIG_SAMPLE);
      sheet.getRange("C2:C" + LINKS_SAMPLE.length).numberFormat = LINKS_SAMPLE.slice(1).map(() => ["0.0"]);
      sheet.getRange("A:N").format.autofitColumns();
      await ctx.sync();
    });
    $("settings").hidden = true;
    await applyTables({ links: created.links, nodes: created.nodes, config: created.config });
    showBanner('Dane przykładowe są w arkuszu „' + created.sheet + '”. Zmieniaj je, a wykres odświeży się sam.');
  }

  function addTable(sheet, topLeft, name, rows) {
    const range = sheet.getRange(topLeft).getResizedRange(rows.length - 1, rows[0].length - 1);
    range.values = rows;
    const table = sheet.tables.add(range, true);
    table.name = name;
    table.style = "TableStyleLight9";
    return table;
  }

  // ---------- UI pomocnicze ----------
  function clearChart() {
    $("chart").innerHTML = "";
    lastSvg = null;
  }

  function showMessage(html, hint, withActions) {
    $("msg-text").innerHTML = html;
    $("msg-hint").textContent = hint || "";
    $("msg-actions").hidden = !withActions || !inExcel;
    $("message").hidden = false;
  }

  function hideMessage() {
    $("message").hidden = true;
  }

  function showBanner(text) {
    const b = $("banner");
    b.textContent = text;
    b.hidden = false;
    clearTimeout(b._t);
    b._t = setTimeout(() => (b.hidden = true), 6000);
  }

  function esc(s) {
    return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
  }

  // soft = true: błąd pokazany w pasku, wykres zostaje (np. nieudane wstawienie kopii)
  async function run(fn, soft) {
    try {
      await fn();
    } catch (e) {
      console.error(e);
      if (soft) {
        showBanner("Błąd: " + (e && e.message ? e.message : String(e)));
        return;
      }
      clearChart();
      showMessage("Błąd: " + esc(e && e.message ? e.message : String(e)), "Sprawdź dane w tabelach albo ustawienia (⚙).", true);
    }
  }

  // Przyciski w komunikacie (stan pusty / błąd)
  document.addEventListener("click", (ev) => {
    const id = ev.target && ev.target.id;
    if (id === "msg-sample") run(createSample);
    if (id === "msg-settings") run(openSettings, true);
  });
})();
