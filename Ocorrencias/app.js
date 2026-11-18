const STORAGE_KEY = "ocorrencias:v1";
const THEME_KEY = "ocorrencias-theme";
const MAX_PASTE_IMAGES = 6;
const MAX_IMAGE_BYTES = 2_000_000; // ~2MB por imagem (base64 aumenta um pouco)

let draftDetalhesImagens = [];
let editDetalhesImagens = [];
let tableSort = { key: "data", dir: "desc" };

function uid() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function todayISO() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function loadAll() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveAll(items) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
}

function normalizeText(v) {
  return String(v || "")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .trim();
}

function cmpByData(a, b, dir = "desc") {
  const aData = a.data || "";
  const bData = b.data || "";
  const byDate = aData.localeCompare(bData);
  if (byDate !== 0) return dir === "asc" ? byDate : -byDate;
  const aUpd = a.atualizadoEm || a.criadoEm || "";
  const bUpd = b.atualizadoEm || b.criadoEm || "";
  return dir === "asc" ? aUpd.localeCompare(bUpd) : bUpd.localeCompare(aUpd);
}

function cmpDateTimeDesc(a, b) {
  return cmpByData(a, b, "desc");
}

function sortRows(rows) {
  const copy = [...rows];
  if (tableSort.key === "data") {
    copy.sort((a, b) => cmpByData(a, b, tableSort.dir));
  }
  return copy;
}

function updateSortHeaderUI() {
  const btn = document.getElementById("btnSortData");
  if (!btn) return;
  const icon = btn.querySelector(".sortIcon");
  const isAsc = tableSort.dir === "asc";
  btn.setAttribute("aria-sort", isAsc ? "ascending" : "descending");
  btn.title = isAsc ? "Ordenar: data mais antiga primeiro" : "Ordenar: data mais recente primeiro";
  if (icon) icon.textContent = isAsc ? "▲" : "▼";
}

function toCSV(rows) {
  const columns = [
    { key: "id", label: "id" },
    { key: "data", label: "data" },
    { key: "severidade", label: "severidade" },
    { key: "status", label: "status" },
    { key: "caderno", label: "caderno" },
    { key: "secao", label: "secao" },
    { key: "resumo", label: "resumo" },
    { key: "responsavel", label: "responsavel" },
    { key: "detalhes", label: "detalhes" },
    { key: "acao", label: "observacoes" },
    { key: "criadoEm", label: "criadoEm" },
    { key: "atualizadoEm", label: "atualizadoEm" },
  ];
  const escape = (v) => {
    const s = String(v ?? "");
    const safe = s.replaceAll('"', '""');
    return `"${safe}"`;
  };
  const lines = [columns.map((c) => c.label).join(",")];
  for (const r of rows) {
    lines.push(
      columns
        .map((c) => escape(r[c.key]))
        .join(","),
    );
  }
  return "\ufeff" + lines.join("\n");
}

function download(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

const EXCEL_COLS = [
  { k: "data", t: "Data", w: 72, style: "Cell" },
  { k: "severidade", t: "Severidade", w: 88, style: "Cell" },
  { k: "status", t: "Status", w: 100, style: "Cell" },
  { k: "caderno", t: "Caderno", w: 95, style: "Cell" },
  { k: "secao", t: "Seção", w: 140, style: "CellWrap" },
  { k: "resumo", t: "Resumo", w: 200, style: "CellWrap" },
  { k: "responsavel", t: "Responsável", w: 110, style: "Cell" },
  { k: "detalhes", t: "Detalhes", w: 420, style: "CellDetalhes" },
  { k: "acao", t: "Observações", w: 220, style: "CellWrap" },
  { k: "imgInfo", t: "Imagens", w: 90, style: "Cell" },
];

function escapeXmlText(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\n/g, "&#10;");
}

function excelSheetName(isoDate) {
  const label = fmtDatePt(isoDate) || "Sem data";
  return label.replace(/[\\/*?:\[\]]/g, "-").slice(0, 31);
}

function rowCellValue(r, key) {
  if (key === "data") return fmtDatePt(r.data) || "";
  if (key === "imgInfo") {
    const n = Array.isArray(r.detalhesImagens) ? r.detalhesImagens.length : 0;
    return n ? `${n} imagem(ns) anexada(s)` : "";
  }
  return String(r[key] ?? "");
}

function buildWorksheetXml(sheetName, dayRows) {
  const headerCells = EXCEL_COLS.map(
    (c) =>
      `<Cell ss:StyleID="Header"><Data ss:Type="String">${escapeXmlText(c.t)}</Data></Cell>`,
  ).join("");

  const colDefs = EXCEL_COLS.map((c) => `<Column ss:Width="${c.w}"/>`).join("");

  const bodyRows = dayRows
    .map((r) => {
      const cells = EXCEL_COLS.map((c) => {
        const val = rowCellValue(r, c.k);
        return `<Cell ss:StyleID="${c.style}"><Data ss:Type="String">${escapeXmlText(val)}</Data></Cell>`;
      }).join("");
      return `<Row ss:AutoFitHeight="1">${cells}</Row>`;
    })
    .join("");

  return `
  <Worksheet ss:Name="${escapeXmlText(sheetName)}">
    <Table>
      ${colDefs}
      <Row ss:StyleID="Header">${headerCells}</Row>
      ${bodyRows}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
      <Selected/>
      <FreezePanes/>
      <FrozenNoSplit/>
      <SplitHorizontal>1</SplitHorizontal>
      <TopRowBottomPane>1</TopRowBottomPane>
      <ActivePane>2</ActivePane>
    </WorksheetOptions>
  </Worksheet>`;
}

function buildExcelWorkbookXml(rows) {
  const byDate = new Map();
  for (const r of rows) {
    const key = r.data || "";
    if (!byDate.has(key)) byDate.set(key, []);
    byDate.get(key).push(r);
  }

  const dates = [...byDate.keys()].sort((a, b) => b.localeCompare(a));
  const worksheets = dates
    .map((date) => {
      const dayRows = byDate.get(date).sort(cmpDateTimeDesc);
      return buildWorksheetXml(excelSheetName(date), dayRows);
    })
    .join("");

  const genAt = escapeXmlText(new Date().toLocaleString("pt-BR"));

  return `\ufeff<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
  <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
    <Title>Relatório de Ocorrências</Title>
    <Created>${genAt}</Created>
  </DocumentProperties>
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal">
      <Alignment ss:Vertical="Top"/>
      <Font ss:FontName="Segoe UI" ss:Size="10"/>
    </Style>
    <Style ss:ID="Header">
      <Font ss:FontName="Segoe UI" ss:Size="10" ss:Bold="1"/>
      <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>
      <Alignment ss:Vertical="Center" ss:Horizontal="Left"/>
    </Style>
    <Style ss:ID="Cell">
      <Alignment ss:Vertical="Top"/>
    </Style>
    <Style ss:ID="CellWrap">
      <Alignment ss:Vertical="Top" ss:WrapText="1"/>
    </Style>
    <Style ss:ID="CellDetalhes">
      <Alignment ss:Vertical="Top" ss:WrapText="1"/>
      <Font ss:FontName="Segoe UI" ss:Size="10"/>
    </Style>
  </Styles>
  ${worksheets}
</Workbook>`;
}

function safePptText(v) {
  const s = String(v ?? "").trim();
  return s.length ? s : "-";
}

function fmtDatePt(iso) {
  if (!iso) return "";
  // iso esperado: YYYY-MM-DD
  const [y, m, d] = String(iso).split("-");
  if (!y || !m || !d) return String(iso);
  return `${d}/${m}/${y}`;
}

function breakLongTokens(text, maxLen = 56) {
  return String(text ?? "")
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map((line) => {
      if (line.length <= maxLen + 10) return line;
      return line.replace(/(\S{58,})/g, (chunk) => {
        const parts = [];
        for (let i = 0; i < chunk.length; i += maxLen) parts.push(chunk.slice(i, i + maxLen));
        return parts.join("\n");
      });
    })
    .join("\n");
}

/** Estimativa conservadora (sem quebrar URL antes da contagem). */
function estimateTextLines(text, charsPerLine = 52) {
  const lines = String(text ?? "").split("\n");
  let total = 0;
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      total += 0.3;
      continue;
    }
    const cpl = /https?:\/\//i.test(trimmed) ? 44 : charsPerLine;
    total += Math.max(1, Math.ceil(trimmed.length / cpl));
  }
  return total;
}

function parseDetalhesBlocks(text) {
  const normalized = String(text ?? "").trim().replace(/\r\n/g, "\n");
  if (!normalized) return [];

  const blocks = [];
  const paragraphs = normalized.split(/\n\s*\n/).filter(Boolean);
  for (const p of paragraphs) {
    const lines = p.split("\n");
    let current = "";
    for (const line of lines) {
      if (/^\d+\s*[-–.]/.test(line.trim()) && current) {
        blocks.push(current.trim());
        current = line;
      } else {
        current = current ? `${current}\n${line}` : line;
      }
    }
    if (current.trim()) blocks.push(current.trim());
  }
  return blocks.length ? blocks : [normalized];
}

function isNumberedItemBlock(block) {
  return /^\d+\s*[-–.]/.test(String(block ?? "").trim());
}

/** Divide blocos em 2 colunas (itens 1–N à esquerda, restante à direita). */
function splitBlocksTwoColumns(blocks) {
  if (!blocks.length) return { col1: "-", col2: "", useTwoCols: false };

  const intro = blocks.filter((b) => !isNumberedItemBlock(b));
  const items = blocks.filter((b) => isNumberedItemBlock(b));

  if (items.length <= 1) {
    const all = blocks.join("\n\n");
    return { col1: all, col2: "", useTwoCols: false };
  }

  const half = Math.ceil(items.length / 2);
  const col1 = [...intro, ...items.slice(0, half)].join("\n\n");
  const col2 = items.slice(half).join("\n\n");
  return { col1: col1 || "-", col2, useTwoCols: !!col2 };
}

/**
 * Só cria 2º slide se realmente não couber.
 * Capacidade: ~26 linhas por coluna (2 colunas) ≈ 52 unidades de linha.
 */
function paginateDetalhesBlocks(blocks, maxLineUnits = 52) {
  if (!blocks.length) return [[]];

  const intro = blocks.filter((b) => !isNumberedItemBlock(b));
  const items = blocks.filter((b) => isNumberedItemBlock(b));
  const introLines = intro.length ? estimateTextLines(intro.join("\n\n")) : 0;
  const totalLines =
    introLines + items.reduce((s, b) => s + estimateTextLines(b), 0);

  if (totalLines <= maxLineUnits) return [blocks];

  const pages = [];
  let pageItems = [];
  let usedLines = introLines;

  for (const item of items) {
    const itemLines = estimateTextLines(item);
    if (pageItems.length && usedLines + itemLines > maxLineUnits) {
      pages.push([...intro, ...pageItems]);
      pageItems = [item];
      usedLines = itemLines;
    } else {
      pageItems.push(item);
      usedLines += itemLines;
    }
  }
  if (pageItems.length) {
    pages.push(pages.length === 0 ? [...intro, ...pageItems] : [...pageItems]);
  }

  return pages.length ? pages : [blocks];
}

function pptTextOpts(extra = {}) {
  return {
    fontFace: "Segoe UI",
    fontSize: 10,
    color: "E8EEFC",
    valign: "top",
    wrap: true,
    lineSpacing: 13,
    ...extra,
  };
}

async function exportPptx(rows) {
  const PptxGenJS = window.PptxGenJS;
  if (!PptxGenJS) {
    alert("Não foi possível carregar o gerador de PPTX. Verifique sua conexão com a internet e tente novamente.");
    return;
  }

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

  const W = 13.33;
  const H = 7.5;
  const margin = 0.5;

  const bg = "0B1220";
  const card = "0F1A30";
  const border = "2A3550";
  const text = "E8EEFC";
  const muted = "A9B7D6";
  const primary = "5B8CFF";
  const danger = "FF4D4D";
  const warn = "FFB020";
  const ok = "2FD57A";

  const now = new Date();
  const genAt = now.toLocaleString("pt-BR");

  const addHeader = (slide, title, subtitle) => {
    slide.background = { color: bg };
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: W,
      h: 1.05,
      fill: { color: card },
      line: { color: border, width: 1 },
    });
    slide.addText(title, {
      x: margin,
      y: 0.22,
      w: W - margin * 2,
      h: 0.4,
      fontFace: "Segoe UI",
      fontSize: 22,
      bold: true,
      color: text,
    });
    slide.addText(subtitle || "", {
      x: margin,
      y: 0.66,
      w: W - margin * 2,
      h: 0.3,
      fontFace: "Segoe UI",
      fontSize: 12,
      color: muted,
    });
  };

  // Capa
  {
    const slide = pptx.addSlide();
    slide.background = { color: bg };
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 1.0,
      y: 1.5,
      w: W - 2.0,
      h: 4.5,
      fill: { color: card },
      line: { color: border, width: 1 },
      radius: 0.2,
    });
    slide.addText("Relatório de Ocorrências", {
      x: 1.2,
      y: 2.15,
      w: W - 2.4,
      h: 0.8,
      fontFace: "Segoe UI",
      fontSize: 34,
      bold: true,
      color: text,
    });
    slide.addText(`Gerado em: ${genAt}`, {
      x: 1.2,
      y: 3.1,
      w: W - 2.4,
      h: 0.4,
      fontFace: "Segoe UI",
      fontSize: 14,
      color: muted,
    });
    slide.addText(`Total no relatório: ${rows.length}`, {
      x: 1.2,
      y: 3.6,
      w: W - 2.4,
      h: 0.4,
      fontFace: "Segoe UI",
      fontSize: 14,
      color: primary,
      bold: true,
    });
    slide.addText("Dica: este PPTX reflete os filtros aplicados na tela.", {
      x: 1.2,
      y: 4.2,
      w: W - 2.4,
      h: 0.4,
      fontFace: "Segoe UI",
      fontSize: 12,
      color: muted,
    });
  }

  const addOccurrenceSlides = (r, idx) => {
    const titulo = `${fmtDatePt(r.data)} • ${safePptText(r.caderno)} / ${safePptText(r.secao)}`;
    const sev = safePptText(r.severidade);
    const status = safePptText(r.status);
    const sevColor = sev === "Crítica" ? warn : sev === "Alta" ? danger : sev === "Baixa" ? ok : primary;
    const stColor = status === "Resolvida" || status === "Encerrada" ? ok : status === "Bloqueada" ? danger : primary;

    const allBlocks = parseDetalhesBlocks(r.detalhes);
    const detailsPages = paginateDetalhesBlocks(allBlocks);
    const imgs = Array.isArray(r.detalhesImagens) ? r.detalhesImagens : [];
    const totalParts = detailsPages.length + (imgs.length ? 1 : 0);

    const drawCardChrome = (slide, partLabel) => {
      const cardX = margin;
      const cardY = 1.25;
      const cardW = W - margin * 2;
      const cardH = H - cardY - margin;

      slide.addShape(pptx.ShapeType.roundRect, {
        x: cardX,
        y: cardY,
        w: cardW,
        h: cardH,
        fill: { color: "0C1629" },
        line: { color: border, width: 1 },
        radius: 0.15,
      });

      const badgeY = cardY + 0.18;
      slide.addShape(pptx.ShapeType.roundRect, {
        x: cardX + 0.25,
        y: badgeY,
        w: 2.1,
        h: 0.42,
        fill: { color: sevColor },
        line: { color: sevColor },
        radius: 0.18,
      });
      slide.addText(`Severidade: ${sev}`, {
        x: cardX + 0.35,
        y: badgeY + 0.09,
        w: 1.9,
        h: 0.28,
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "0B1220",
        bold: true,
      });

      slide.addShape(pptx.ShapeType.roundRect, {
        x: cardX + 2.45,
        y: badgeY,
        w: 2.15,
        h: 0.42,
        fill: { color: stColor },
        line: { color: stColor },
        radius: 0.18,
      });
      slide.addText(`Status: ${status}`, {
        x: cardX + 2.55,
        y: badgeY + 0.09,
        w: 1.95,
        h: 0.28,
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "0B1220",
        bold: true,
      });

      slide.addText(`Responsável: ${safePptText(r.responsavel)}`, {
        x: cardX + 4.75,
        y: badgeY + 0.07,
        w: cardW - 5.0,
        h: 0.32,
        fontFace: "Segoe UI",
        fontSize: 11,
        color: muted,
      });

      if (partLabel) {
        slide.addText(partLabel, {
          x: cardX + cardW - 2.6,
          y: badgeY + 0.07,
          w: 2.35,
          h: 0.32,
          fontFace: "Segoe UI",
          fontSize: 10,
          color: primary,
          align: "right",
        });
      }

      return { cardX, cardY, cardW, cardH };
    };

    const drawResumoObs = (slide, geom) => {
      const pad = 0.25;
      const x = geom.cardX + pad;
      const w = geom.cardW - pad * 2;
      let y = geom.cardY + 0.68;
      const obs = safePptText(r.acao);
      const hasObs = obs !== "-";

      slide.addText("Resumo", {
        x,
        y,
        w,
        h: 0.26,
        fontFace: "Segoe UI",
        fontSize: 13,
        bold: true,
        color: text,
      });
      slide.addText(safePptText(r.resumo), {
        x,
        y: y + 0.28,
        w,
        h: 0.42,
        ...pptTextOpts({ fontSize: 11 }),
      });
      y += 0.78;

      if (hasObs) {
        slide.addText("Observações", {
          x,
          y,
          w,
          h: 0.26,
          fontFace: "Segoe UI",
          fontSize: 13,
          bold: true,
          color: text,
        });
        slide.addText(obs, {
          x,
          y: y + 0.28,
          w,
          h: 0.42,
          ...pptTextOpts({ fontSize: 10 }),
        });
        y += 0.78;
      }

      return y;
    };

    const drawDetalhesColumns = (slide, geom, pageBlocks, opts = {}) => {
      const pad = 0.25;
      const x = geom.cardX + pad;
      const w = geom.cardW - pad * 2;
      const label = opts.continuation ? "Detalhes (continuação)" : "Detalhes";
      const bodyTop = opts.bodyTop ?? geom.cardY + 1.48;
      const bottom = geom.cardY + geom.cardH - 0.2;
      const boxY = bodyTop;
      const boxH = Math.max(1.5, bottom - boxY);

      slide.addText(label, {
        x,
        y: boxY - 0.3,
        w,
        h: 0.26,
        fontFace: "Segoe UI",
        fontSize: 13,
        bold: true,
        color: text,
      });

      const layout = splitBlocksTwoColumns(pageBlocks);
      const gap = 0.22;
      const col1Text = breakLongTokens(layout.col1);
      const col2Text = breakLongTokens(layout.col2);

      if (!layout.useTwoCols) {
        slide.addText(col1Text, {
          x,
          y: boxY,
          w,
          h: boxH,
          ...pptTextOpts(),
        });
        return;
      }

      const colW = (w - gap) / 2;
      slide.addText(col1Text, {
        x,
        y: boxY,
        w: colW,
        h: boxH,
        ...pptTextOpts(),
      });
      slide.addText(col2Text, {
        x: x + colW + gap,
        y: boxY,
        w: colW,
        h: boxH,
        ...pptTextOpts(),
      });
    };

    detailsPages.forEach((pageBlocks, pageIdx) => {
      const slide = pptx.addSlide();
      const partNum = pageIdx + 1;
      const partLabel =
        totalParts > 1 ? `Parte ${partNum}/${totalParts}` : "";
      const isContinuation = pageIdx > 0;
      addHeader(
        slide,
        `Ocorrência ${idx + 1} de ${rows.length}`,
        titulo + (detailsPages.length > 1 ? ` • Detalhes (${partNum}/${detailsPages.length})` : ""),
      );
      const geom = drawCardChrome(slide, partLabel);

      let detalhesTop;
      if (isContinuation) {
        detalhesTop = geom.cardY + 0.72;
      } else {
        const afterResumo = drawResumoObs(slide, geom);
        detalhesTop = Math.max(afterResumo + 0.12, geom.cardY + 1.48);
      }

      drawDetalhesColumns(slide, geom, pageBlocks, {
        continuation: isContinuation,
        bodyTop: detalhesTop,
      });
    });

    if (imgs.length) {
      const slide = pptx.addSlide();
      const partLabel = totalParts > 1 ? `Parte ${totalParts}/${totalParts}` : "";
      addHeader(slide, `Ocorrência ${idx + 1} de ${rows.length}`, `${titulo} • Imagens`);
      const geom = drawCardChrome(slide, partLabel);

      const pad = 0.25;
      const x = geom.cardX + pad;
      const y = geom.cardY + 0.72;
      const w = geom.cardW - pad * 2;
      const h = geom.cardH - 1.1;

      slide.addText("Imagens anexadas", {
        x,
        y: y - 0.05,
        w,
        h: 0.28,
        fontFace: "Segoe UI",
        fontSize: 13,
        bold: true,
        color: text,
      });

      const imgMax = Math.min(2, imgs.length);
      const gap = 0.2;
      const imgW = (w - gap * (imgMax - 1)) / imgMax;
      for (let i = 0; i < imgMax; i++) {
        try {
          slide.addImage({
            data: imgs[i],
            x: x + i * (imgW + gap),
            y,
            w: imgW,
            h,
          });
        } catch {
          // ignora imagem inválida
        }
      }
    }
  };

  rows.forEach((r, idx) => addOccurrenceSlides(r, idx));

  const ts = new Date().toISOString().slice(0, 19).replaceAll(":", "-");
  await pptx.writeFile({ fileName: `ocorrencias-${ts}.pptx` });
}

function el(id) {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Elemento não encontrado: ${id}`);
  return node;
}

function renderPillSev(sev) {
  const s = String(sev || "");
  const span = document.createElement("span");
  span.className = "pill";
  span.textContent = s;
  if (s === "Crítica") span.classList.add("pill--crit");
  if (s === "Alta") span.classList.add("pill--danger");
  if (s === "Baixa") span.classList.add("pill--ok");
  return span;
}

function renderPillStatus(status) {
  const s = String(status || "");
  const span = document.createElement("span");
  span.className = "pill";
  span.textContent = s;
  if (s === "Resolvida" || s === "Encerrada") span.classList.add("pill--ok");
  if (s === "Bloqueada") span.classList.add("pill--danger");
  return span;
}

function applyFilters(items, f) {
  const de = f.de ? `${f.de}T00:00` : null;
  const ate = f.ate ? `${f.ate}T23:59` : null;
  const busca = normalizeText(f.busca);
  const somenteAbertas = !!f.somenteAbertas;

  return items.filter((it) => {
    const dt = `${it.data || ""}T00:00`;
    if (de && dt < de) return false;
    if (ate && dt > ate) return false;
    if (somenteAbertas) {
      const isOpen = it.status === "Aberta" || it.status === "Em andamento" || it.status === "Bloqueada";
      if (!isOpen) return false;
    } else if (f.status.length && !f.status.includes(it.status)) {
      return false;
    }
    if (f.severidade.length && !f.severidade.includes(it.severidade)) return false;

    if (busca) {
      const blob = normalizeText(
        [
          it.data,
          it.severidade,
          it.status,
          it.caderno,
          it.secao,
          it.resumo,
          it.responsavel,
          it.detalhes,
          it.acao,
        ].join(" "),
      );
      if (!blob.includes(busca)) return false;
    }
    return true;
  });
}

function isCadernoOutros(caderno) {
  return caderno === "Outros";
}

function sectionsFor(caderno) {
  switch (caderno) {
    case "Executivo":
      return ["Atos Normativos", "Atos de Pessoal", "Atos de Gestão e Despesas"];
    case "Legislativo":
      return [
        "Atos Legislativos e Parlamentares de Assembléia",
        "Atos Administrativos da Assembleia e Tribunal de Contas",
      ];
    case "Municípios":
      return ["Atos municipais"];
    case "Empresarial":
      return ["Atos Empresariais", "Atos da Jucesp"];
    case "Outros":
      return [];
    default:
      return [];
  }
}

function readSecaoValue(isEdit = false) {
  const caderno = el(isEdit ? "eCaderno" : "caderno").value;
  if (isCadernoOutros(caderno)) {
    return el(isEdit ? "eSecaoText" : "secaoText").value.trim();
  }
  return el(isEdit ? "eSecao" : "secao").value;
}

function updateSecaoField(isEdit, caderno, selectedValue = "") {
  const selectEl = el(isEdit ? "eSecao" : "secao");
  const textEl = el(isEdit ? "eSecaoText" : "secaoText");
  const labelEl = el(isEdit ? "eSecaoLabel" : "secaoLabel");

  if (isCadernoOutros(caderno)) {
    selectEl.hidden = true;
    selectEl.required = false;
    selectEl.disabled = true;
    textEl.hidden = false;
    textEl.required = true;
    textEl.disabled = false;
    textEl.value = selectedValue || "";
    labelEl.textContent = "Local da ocorrência";
    return;
  }

  selectEl.hidden = false;
  selectEl.required = true;
  selectEl.disabled = false;
  textEl.hidden = true;
  textEl.required = false;
  textEl.disabled = true;
  textEl.value = "";
  labelEl.textContent = "Seção";
  populateSecaoSelect(selectEl, caderno, selectedValue);
}

function populateSecaoSelect(selectEl, caderno, selectedValue = "") {
  const options = sectionsFor(caderno);
  selectEl.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.disabled = true;
  placeholder.selected = !selectedValue;
  placeholder.textContent = caderno ? "Selecione" : "Selecione o caderno primeiro";
  selectEl.appendChild(placeholder);

  for (const opt of options) {
    const o = document.createElement("option");
    o.value = opt;
    o.textContent = opt;
    if (selectedValue && selectedValue === opt) o.selected = true;
    selectEl.appendChild(o);
  }

  if (selectedValue && !options.includes(selectedValue)) {
    selectEl.value = "";
  }
}

function computeStats(items) {
  const total = items.length;
  const open = items.filter((i) => i.status === "Aberta" || i.status === "Em andamento" || i.status === "Bloqueada").length;
  const critical = items.filter((i) => i.severidade === "Crítica").length;
  return { total, open, critical };
}

function getCheckedInGroup(groupId) {
  return Array.from(el(groupId).querySelectorAll('input[type="checkbox"]:checked')).map((cb) => cb.value);
}

function clearCheckedInGroup(groupId) {
  el(groupId).querySelectorAll('input[type="checkbox"]').forEach((cb) => {
    cb.checked = false;
  });
}

function getFiltersFromUI() {
  return {
    de: el("fDe").value || "",
    ate: el("fAte").value || "",
    somenteAbertas: !!el("fSomenteAbertas").checked,
    status: getCheckedInGroup("fStatusGroup"),
    severidade: getCheckedInGroup("fSeveridadeGroup"),
    busca: el("fBusca").value || "",
  };
}

function setDefaultDatesIfEmpty() {
  const data = el("data");
  if (!data.value) data.value = todayISO();
}

function renderTable(rows) {
  const tbody = el("tbody");
  const empty = el("emptyState");

  tbody.innerHTML = "";

  if (rows.length === 0) {
    empty.hidden = false;
    return;
  }
  empty.hidden = true;

  for (const r of rows) {
    const tr = document.createElement("tr");

    const tdData = document.createElement("td");
    tdData.textContent = fmtDatePt(r.data) || r.data || "";

    const tdSev = document.createElement("td");
    tdSev.appendChild(renderPillSev(r.severidade));

    const tdStatus = document.createElement("td");
    tdStatus.appendChild(renderPillStatus(r.status));

    const tdCaderno = document.createElement("td");
    tdCaderno.textContent = r.caderno || "";

    const tdSecao = document.createElement("td");
    tdSecao.textContent = r.secao || "";

    const tdResumo = document.createElement("td");
    tdResumo.textContent = r.resumo || "";

    const tdResp = document.createElement("td");
    tdResp.textContent = r.responsavel || "";

    const tdAcoes = document.createElement("td");
    tdAcoes.className = "noPrint";
    const wrap = document.createElement("div");
    wrap.className = "rowActions";

    const btnEdit = document.createElement("button");
    btnEdit.type = "button";
    btnEdit.className = "linkBtn";
    btnEdit.textContent = "Editar";
    btnEdit.addEventListener("click", () => openEdit(r.id));

    const btnDel = document.createElement("button");
    btnDel.type = "button";
    btnDel.className = "linkBtn";
    btnDel.textContent = "Excluir";
    btnDel.addEventListener("click", () => removeOne(r.id));

    wrap.append(btnEdit, btnDel);
    tdAcoes.appendChild(wrap);

    tr.append(tdData, tdSev, tdStatus, tdCaderno, tdSecao, tdResumo, tdResp, tdAcoes);
    tbody.appendChild(tr);
  }
}

function renderImageStrip(containerEl, images, onRemove) {
  containerEl.innerHTML = "";
  if (!Array.isArray(images) || images.length === 0) return;

  images.forEach((src, idx) => {
    const thumb = document.createElement("div");
    thumb.className = "imgThumb";

    const img = document.createElement("img");
    img.src = src;
    img.alt = `Imagem ${idx + 1}`;

    const btn = document.createElement("button");
    btn.type = "button";
    btn.title = "Remover imagem";
    btn.textContent = "×";
    btn.addEventListener("click", () => onRemove(idx));

    thumb.append(img, btn);
    containerEl.appendChild(thumb);
  });
}

async function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Falha ao ler a imagem"));
    reader.onload = () => resolve(String(reader.result || ""));
    reader.readAsDataURL(file);
  });
}

async function handlePasteImages(ev, getImages, setImages, renderFn) {
  const dt = ev.clipboardData;
  if (!dt?.items?.length) return;

  const imageItems = Array.from(dt.items).filter((it) => it.kind === "file" && /^image\//.test(it.type));
  if (imageItems.length === 0) return;

  ev.preventDefault();

  const current = Array.isArray(getImages()) ? [...getImages()] : [];
  const spaceLeft = Math.max(0, MAX_PASTE_IMAGES - current.length);
  const toTake = imageItems.slice(0, spaceLeft);

  if (spaceLeft === 0) {
    alert(`Limite de ${MAX_PASTE_IMAGES} imagens por ocorrência atingido.`);
    return;
  }

  for (const item of toTake) {
    const file = item.getAsFile();
    if (!file) continue;
    if (file.size > MAX_IMAGE_BYTES) {
      alert("Uma imagem colada é grande demais. Use imagens menores (até ~2MB).");
      continue;
    }
    const dataUrl = await fileToDataUrl(file);
    current.push(dataUrl);
  }

  setImages(current);
  renderFn();
}

function refresh() {
  const all = loadAll();
  const filters = getFiltersFromUI();
  const filtered = sortRows(applyFilters(all, filters));

  const stats = computeStats(filtered);
  el("totalCount").textContent = String(stats.total);
  el("openCount").textContent = String(stats.open);
  el("criticalCount").textContent = String(stats.critical);

  updateSortHeaderUI();
  renderTable(filtered);

  // Export sempre respeita filtro atual
  window.__ocorrenciasFiltered = filtered;
}

function getFormData(prefix = "") {
  const get = (id) => el(prefix + id).value?.trim?.() ?? el(prefix + id).value ?? "";
  return {
    data: get("data"),
    severidade: get("severidade"),
    status: get("status"),
    responsavel: get("responsavel"),
    caderno: get("caderno"),
    secao: readSecaoValue(!!prefix),
    resumo: get("resumo"),
    detalhes: get("detalhes"),
    detalhesImagens: prefix ? editDetalhesImagens : draftDetalhesImagens,
    acao: get("acao"),
  };
}

function validateRequired(payload) {
  const required = ["data", "severidade", "status", "caderno", "secao", "resumo"];
  for (const k of required) {
    if (!payload[k]) return k;
  }
  return "";
}

function addOne(payload) {
  const all = loadAll();
  const now = new Date().toISOString();
  all.push({
    id: uid(),
    ...payload,
    criadoEm: now,
    atualizadoEm: now,
  });
  saveAll(all);
}

function updateOne(id, changes) {
  const all = loadAll();
  const idx = all.findIndex((x) => x.id === id);
  if (idx < 0) return;
  all[idx] = { ...all[idx], ...changes, atualizadoEm: new Date().toISOString() };
  saveAll(all);
}

function removeOne(id) {
  const all = loadAll();
  const item = all.find((x) => x.id === id);
  if (!item) return;
  const ok = confirm(`Excluir a ocorrência de ${item.data || ""} - "${item.resumo || ""}"?`);
  if (!ok) return;
  saveAll(all.filter((x) => x.id !== id));
  refresh();
}

function clearAll() {
  const ok = confirm("Isso vai apagar TODAS as ocorrências deste navegador. Deseja continuar?");
  if (!ok) return;
  localStorage.removeItem(STORAGE_KEY);
  refresh();
}

function openEdit(id) {
  const all = loadAll();
  const it = all.find((x) => x.id === id);
  if (!it) return;

  editDetalhesImagens = Array.isArray(it.detalhesImagens) ? [...it.detalhesImagens] : [];

  el("editId").value = it.id;
  el("eData").value = it.data || "";
  el("eSeveridade").value = it.severidade || "Baixa";
  el("eStatus").value = it.status || "Aberta";
  el("eResponsavel").value = it.responsavel || "";
  el("eCaderno").value = it.caderno || "Executivo";
  updateSecaoField(true, el("eCaderno").value, it.secao || "");
  el("eResumo").value = it.resumo || "";
  el("eDetalhes").value = it.detalhes || "";
  renderImageStrip(el("eDetalhesImgs"), editDetalhesImagens, (idx) => {
    editDetalhesImagens.splice(idx, 1);
    renderImageStrip(el("eDetalhesImgs"), editDetalhesImagens, (i) => {
      editDetalhesImagens.splice(i, 1);
      renderImageStrip(el("eDetalhesImgs"), editDetalhesImagens, () => {});
    });
  });
  el("eAcao").value = it.acao || "";

  const dlg = el("dlgEdit");
  dlg.showModal();
}

function initTheme() {
  const btn = document.getElementById("btnToggleTheme");
  const saved = localStorage.getItem(THEME_KEY);
  const theme = saved === "light" ? "light" : "dark";
  document.documentElement.setAttribute("data-theme", theme);
  if (btn) btn.textContent = theme === "light" ? "Tema escuro" : "Tema claro";
  if (!btn) return;
  btn.addEventListener("click", () => {
    const current = document.documentElement.getAttribute("data-theme") === "light" ? "light" : "dark";
    const next = current === "light" ? "dark" : "light";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem(THEME_KEY, next);
    btn.textContent = next === "light" ? "Tema escuro" : "Tema claro";
  });
}

function init() {
  initTheme();
  setDefaultDatesIfEmpty();
  updateSecaoField(false, el("caderno").value || "", "");
  el("caderno").addEventListener("change", () => {
    updateSecaoField(false, el("caderno").value || "", "");
  });

  // Colar imagens no Detalhes (cadastro)
  draftDetalhesImagens = [];
  const renderDraftImages = () =>
    renderImageStrip(el("detalhesImgs"), draftDetalhesImagens, (idx) => {
      draftDetalhesImagens.splice(idx, 1);
      renderDraftImages();
    });
  renderDraftImages();
  el("detalhes").addEventListener("paste", (ev) =>
    handlePasteImages(
      ev,
      () => draftDetalhesImagens,
      (imgs) => {
        draftDetalhesImagens = imgs;
      },
      renderDraftImages,
    ),
  );

  el("formOcorrencia").addEventListener("submit", (ev) => {
    ev.preventDefault();
    const payload = getFormData("");
    const missing = validateRequired(payload);
    if (missing) {
      alert("Preencha os campos obrigatórios antes de salvar.");
      return;
    }
    addOne(payload);
    ev.target.reset();
    setDefaultDatesIfEmpty();
    updateSecaoField(false, el("caderno").value || "", "");
    draftDetalhesImagens = [];
    renderDraftImages();
    refresh();
  });

  el("btnSortData").addEventListener("click", () => {
    tableSort.dir = tableSort.dir === "desc" ? "asc" : "desc";
    refresh();
  });

  el("btnAplicarFiltros").addEventListener("click", refresh);
  el("btnLimparFiltros").addEventListener("click", () => {
    el("fDe").value = "";
    el("fAte").value = "";
    el("fSomenteAbertas").checked = false;
    clearCheckedInGroup("fStatusGroup");
    clearCheckedInGroup("fSeveridadeGroup");
    el("fBusca").value = "";
    refresh();
  });

  el("btnRelatorioAbertas").addEventListener("click", () => {
    el("fSomenteAbertas").checked = true;
    clearCheckedInGroup("fStatusGroup");
    refresh();
    window.print();
  });

  el("btnPrint").addEventListener("click", () => window.print());

  el("btnExportCsv").addEventListener("click", () => {
    const rows = Array.isArray(window.__ocorrenciasFiltered) ? window.__ocorrenciasFiltered : loadAll();
    const csv = toCSV(rows);
    const ts = new Date().toISOString().slice(0, 19).replaceAll(":", "-");
    download(`ocorrencias-${ts}.csv`, csv, "text/csv;charset=utf-8");
  });

  el("btnExportXls").addEventListener("click", () => {
    const rows = Array.isArray(window.__ocorrenciasFiltered) ? window.__ocorrenciasFiltered : loadAll();
    if (!rows.length) {
      alert("Não há ocorrências para exportar (verifique os filtros).");
      return;
    }
    const xml = buildExcelWorkbookXml(rows);
    const ts = new Date().toISOString().slice(0, 19).replaceAll(":", "-");
    download(`ocorrencias-${ts}.xls`, xml, "application/vnd.ms-excel;charset=utf-8");
  });

  el("btnExportPptx").addEventListener("click", async () => {
    const rows = Array.isArray(window.__ocorrenciasFiltered) ? window.__ocorrenciasFiltered : loadAll();
    if (!rows.length) {
      alert("Não há ocorrências para exportar (verifique os filtros).");
      return;
    }
    try {
      await exportPptx(rows);
    } catch (e) {
      alert("Falha ao gerar o PPTX. Tente novamente.");
      console.error(e);
    }
  });

  el("btnLimparTudo").addEventListener("click", clearAll);

  el("btnCancelEdit").addEventListener("click", () => el("dlgEdit").close());
  el("eCaderno").addEventListener("change", () => {
    updateSecaoField(true, el("eCaderno").value || "", "");
  });

  // Colar imagens no Detalhes (edição)
  const renderEditImages = () =>
    renderImageStrip(el("eDetalhesImgs"), editDetalhesImagens, (idx) => {
      editDetalhesImagens.splice(idx, 1);
      renderEditImages();
    });
  el("eDetalhes").addEventListener("paste", (ev) =>
    handlePasteImages(
      ev,
      () => editDetalhesImagens,
      (imgs) => {
        editDetalhesImagens = imgs;
      },
      renderEditImages,
    ),
  );

  el("formEdit").addEventListener("submit", (ev) => {
    ev.preventDefault();
    const id = el("editId").value;
    const payload = {
      data: el("eData").value,
      severidade: el("eSeveridade").value,
      status: el("eStatus").value,
      responsavel: el("eResponsavel").value.trim(),
      caderno: el("eCaderno").value,
      secao: readSecaoValue(true),
      resumo: el("eResumo").value.trim(),
      detalhes: el("eDetalhes").value,
      detalhesImagens: editDetalhesImagens,
      acao: el("eAcao").value,
    };
    const missing = validateRequired(payload);
    if (missing) {
      alert("Preencha os campos obrigatórios antes de salvar.");
      return;
    }
    updateOne(id, payload);
    el("dlgEdit").close();
    refresh();
  });

  // Atualiza automaticamente quando mexer nos filtros
  for (const id of ["fDe", "fAte", "fSomenteAbertas", "fBusca"]) {
    el(id).addEventListener("change", refresh);
    el(id).addEventListener("keyup", (e) => {
      if (e.key === "Enter") refresh();
    });
  }
  for (const id of ["fStatusGroup", "fSeveridadeGroup"]) {
    el(id).addEventListener("change", refresh);
  }

  refresh();
}

document.addEventListener("DOMContentLoaded", init);

