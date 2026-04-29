/* =========================================================
 *  AI Assistant Pro — taskpane.js  v2.0
 *  ИП Посаднев
 *
 *  Принимает Action Plan от бэкенда и исполняет его через Office.js:
 *  set_values, format, borders, formulas, charts, conditional formatting,
 *  freeze panes, sorting, validation, comments, hyperlinks и т.д.
 * ========================================================= */

const API_URL = "https://excel-ai-pro.ru/api/analyze";
const HEALTH_URL = "https://excel-ai-pro.ru/api/health";
const LICENSE_URL = "https://excel-ai-pro.ru/api/license/check";

const STORAGE_KEYS = {
  LICENSE: "aiapro_license",
  HISTORY: "aiapro_history",
};

const CURRENT_MODE = "auto"; // фиксированный режим (UI режимов убран)

/* ---------- ВСПОМОГАТЕЛЬНОЕ ---------- */

function $(id) { return document.getElementById(id); }

function setStatus(cls, text) {
  const r = $("result");
  r.className = cls;
  r.innerText = text;
}

function loadHistory() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEYS.HISTORY) || "[]"); }
  catch { return []; }
}

function saveHistory(item) {
  const arr = loadHistory();
  arr.unshift(item);
  localStorage.setItem(STORAGE_KEYS.HISTORY, JSON.stringify(arr.slice(0, 20)));
  renderHistory();
}

function renderHistory() {
  const list = $("historyList");
  if (!list) return;
  const items = loadHistory();
  list.innerHTML = "";
  if (!items.length) {
    list.innerHTML = '<div class="muted">Здесь появятся ваши последние запросы.</div>';
    return;
  }
  items.forEach((it) => {
    const d = document.createElement("div");
    d.className = "history-item";
    d.title = "Кликните, чтобы вставить запрос";
    d.innerHTML = `<span class="hi-icon">${it.ok ? "✅" : "⚠️"}</span><span class="hi-text">${escapeHtml(it.prompt)}</span>`;
    d.onclick = () => { $("promptInput").value = it.prompt; };
    list.appendChild(d);
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  })[c]);
}

/* ---------- ЛИЦЕНЗИЯ ---------- */

async function verifyLicense() {
  const pin = $("licenseKey").value.trim();
  if (!pin) { setStatus("status-error", "❌ Введите лицензионный ключ."); return; }
  setStatus("status-loading", "⏳ Проверка лицензии...");
  try {
    const res = await fetch(LICENSE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pin }),
    });
    const j = await res.json();
    if (j.valid) {
      localStorage.setItem(STORAGE_KEYS.LICENSE, pin);
      const parts = [`✅ ${j.name}`];
      if (j.limit_total) {
        parts.push(`осталось ${j.remaining_total} из ${j.limit_total}`);
      } else if (j.limit_per_day) {
        parts.push(`сегодня ${j.used_today}/${j.limit_per_day}`);
      } else {
        parts.push("без лимита");
      }
      setStatus("status-success", parts.join(" · "));
    } else {
      setStatus("status-error", "⚠️ " + (j.message || "Лицензия не найдена."));
    }
  } catch (e) {
    setStatus("status-error", "❌ Сервер недоступен: " + e.message);
  }
}

/* ---------- ИСПОЛНИТЕЛЬ ACTION PLAN ---------- */

// Excel: имя листа ≤ 31 символа, нельзя : \ / ? * [ ],
// нельзя начинать/заканчивать апострофом, "History" зарезервировано.
function sanitizeSheetName(name) {
  if (!name) return "Результат ИИ";
  let s = String(name).replace(/[:\\\/\?\*\[\]]/g, "_").trim();
  s = s.replace(/^'+|'+$/g, "");
  if (s.toLowerCase() === "history") s = "История";
  if (s.length > 31) s = s.slice(0, 31);
  return s || "Лист";
}

async function getUniqueSheetName(context, base) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  const existing = new Set(sheets.items.map((s) => s.name));
  let cand = sanitizeSheetName(base);
  if (!existing.has(cand)) return cand;
  for (let i = 2; i < 1000; i++) {
    cand = sanitizeSheetName(`${base} ${i}`);
    if (!existing.has(cand)) return cand;
  }
  return sanitizeSheetName(`${base} ${Date.now()}`);
}

// КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ: используем getItemOrNullObject + load + sync,
// иначе Office.js падает с "Запрошенный ресурс не существует" на следующем sync.
async function resolveSheet(context, action, createdSheets) {
  const wb = context.workbook;
  const rawName = action.sheet || "_active_";

  if (rawName === "_active_") {
    return wb.worksheets.getActiveWorksheet();
  }

  if (rawName === "_new_") {
    const desired = sanitizeSheetName(action.sheet_name || "Результат ИИ");
    const unique = await getUniqueSheetName(context, desired);
    const s = wb.worksheets.add(unique);
    await context.sync(); // КРИТИЧНО: коммитим создание листа перед использованием
    createdSheets.set("_new_", s);
    action.sheet = unique;
    return s;
  }

  const safe = sanitizeSheetName(rawName);
  const maybe = wb.worksheets.getItemOrNullObject(safe);
  maybe.load("isNullObject,name");
  await context.sync();
  if (!maybe.isNullObject) {
    action.sheet = maybe.name;
    return maybe;
  }
  const created = wb.worksheets.add(safe);
  await context.sync(); // КРИТИЧНО: коммитим создание листа
  action.sheet = safe;
  return created;
}

function colLetter(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s || "A";
}

async function executePlan(context, actions) {
  const createdSheets = new Map();

  // На время большого плана отключаем автопересчёт формул (5-10× быстрее)
  let prevCalcMode = null;
  try {
    context.application.load("calculationMode");
    await context.sync();
    prevCalcMode = context.application.calculationMode;
    context.application.calculationMode = "Manual";
    await context.sync();
  } catch (_) { /* старые версии Excel могут не поддерживать */ }

  try {
    for (let i = 0; i < actions.length; i++) {
      const a = actions[i];
      try {
        await runAction(context, a, createdSheets);
        await context.sync();
      } catch (e) {
        console.error("Ошибка действия", a, e);
        throw new Error(`Действие #${i + 1} (${a.type}): ${e.message}`);
      }
    }
  } finally {
    // Возвращаем автопересчёт
    if (prevCalcMode) {
      try {
        context.application.calculationMode = prevCalcMode;
        context.application.calculate("Full");
        await context.sync();
      } catch (_) {}
    }
  }
}

async function runAction(context, a, createdSheets) {
  switch (a.type) {

    case "create_sheet": {
      const safe = sanitizeSheetName(a.name);
      const maybe = context.workbook.worksheets.getItemOrNullObject(safe);
      maybe.load("isNullObject");
      await context.sync();
      let sheet;
      if (maybe.isNullObject) {
        sheet = context.workbook.worksheets.add(safe);
        await context.sync(); // КРИТИЧНО: коммитим, иначе tabColor/activate упадут
      } else {
        sheet = maybe;
      }
      if (a.color) sheet.tabColor = a.color;
      if (a.activate !== false) sheet.activate();
      return;
    }

    case "rename_sheet": {
      const maybe = context.workbook.worksheets.getItemOrNullObject(a.old_name);
      maybe.load("isNullObject");
      await context.sync();
      if (!maybe.isNullObject) maybe.name = sanitizeSheetName(a.new_name);
      return;
    }

    case "delete_sheet": {
      const maybe = context.workbook.worksheets.getItemOrNullObject(a.name);
      maybe.load("isNullObject");
      await context.sync();
      if (!maybe.isNullObject) maybe.delete();
      return;
    }

    case "activate_sheet": {
      const maybe = context.workbook.worksheets.getItemOrNullObject(a.name);
      maybe.load("isNullObject");
      await context.sync();
      if (!maybe.isNullObject) maybe.activate();
      return;
    }

    case "set_values": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const values = a.values || [];
      if (!values.length) return;
      const cols = values[0].length;
      const startCell = a.start_cell || "A1";

      // Получаем индексы стартовой ячейки
      const start = sheet.getRange(startCell);
      start.load("rowIndex,columnIndex");
      await context.sync();

      // Запись чанками для больших объёмов (5000+ строк)
      const CHUNK = values.length > 1000 ? 500 : values.length;
      for (let i = 0; i < values.length; i += CHUNK) {
        const chunk = values.slice(i, i + CHUNK);
        const range = sheet.getRangeByIndexes(
          start.rowIndex + i,
          start.columnIndex,
          chunk.length,
          cols
        );
        range.values = chunk;
        if (CHUNK < values.length) {
          // Промежуточный sync для стабильности и прогресса
          await context.sync();
          if (typeof setStatus === "function") {
            const done = Math.min(i + CHUNK, values.length);
            setStatus("status-loading", `🛠 Запись ${done}/${values.length} строк...`);
          }
        }
      }
      return;
    }

    case "set_formula": {
      const sheet = await resolveSheet(context, a, createdSheets);
      sheet.getRange(a.cell).formulas = [[a.formula]];
      return;
    }

    case "fill_formula": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const range = sheet.getRange(a.range);
      range.load("rowCount,columnCount");
      await context.sync();
      const arr = [];
      for (let r = 0; r < range.rowCount; r++) {
        const row = [];
        for (let c = 0; c < range.columnCount; c++) row.push(a.formula);
        arr.push(row);
      }
      range.formulas = arr;
      return;
    }

    case "format": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const r = sheet.getRange(a.range);
      const f = r.format;
      if (a.fill_color) f.fill.color = a.fill_color;
      if (a.font_color) f.font.color = a.font_color;
      if (a.font_size)  f.font.size = a.font_size;
      if (a.font_name)  f.font.name = a.font_name;
      if (typeof a.bold === "boolean") f.font.bold = a.bold;
      if (typeof a.italic === "boolean") f.font.italic = a.italic;
      if (typeof a.underline === "boolean") f.font.underline = a.underline ? "Single" : "None";
      if (a.horizontal_alignment) f.horizontalAlignment = a.horizontal_alignment;
      if (a.vertical_alignment)   f.verticalAlignment = a.vertical_alignment;
      if (typeof a.wrap_text === "boolean") f.wrapText = a.wrap_text;
      if (a.number_format) r.numberFormat = [[a.number_format]];
      return;
    }

    case "borders": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const r = sheet.getRange(a.range);
      const edges = a.edges || ["EdgeTop","EdgeBottom","EdgeLeft","EdgeRight","InsideHorizontal","InsideVertical"];
      edges.forEach((edge) => {
        const b = r.format.borders.getItem(edge);
        b.style = a.style || "Continuous";
        if (a.weight) b.weight = a.weight;
        if (a.color)  b.color = a.color;
      });
      return;
    }

    case "autofit": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const ur = sheet.getUsedRange();
      if (a.mode === "rows") ur.format.autofitRows();
      else if (a.mode === "both") { ur.format.autofitColumns(); ur.format.autofitRows(); }
      else ur.format.autofitColumns();
      return;
    }

    case "merge_cells": {
      const sheet = await resolveSheet(context, a, createdSheets);
      sheet.getRange(a.range).merge(!!a.across); return;
    }

    case "freeze_panes": {
      const sheet = await resolveSheet(context, a, createdSheets);
      sheet.freezePanes.unfreeze();
      if (a.rows && !a.columns) sheet.freezePanes.freezeRows(a.rows);
      else if (!a.rows && a.columns) sheet.freezePanes.freezeColumns(a.columns);
      else if (a.rows && a.columns) {
        const cell = `${colLetter(a.columns + 1)}${a.rows + 1}`;
        sheet.freezePanes.freezeAt(sheet.getRange(`A1:${cell}`));
      }
      return;
    }

    case "conditional_format": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const r = sheet.getRange(a.range);
      switch (a.rule) {
        case "color_scale": {
          const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
          cf.colorScale.criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: a.min_color || "#FFFFFF" },
            midpoint:{ type: Excel.ConditionalFormatColorCriterionType.percentile, formula: "50", color: a.mid_color || "#FFEB84" },
            maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: a.max_color || "#FF6B6B" },
          };
          break;
        }
        case "data_bars": {
          const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
          cf.dataBar.barColor = a.fill_color || "#4472C4";
          break;
        }
        case "greater_than":
        case "less_than":
        case "between": {
          const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
          const op = a.rule === "greater_than" ? "GreaterThan"
                   : a.rule === "less_than"    ? "LessThan"
                   : "Between";
          cf.cellValue.rule = a.rule === "between"
            ? { formula1: String(a.value), formula2: String(a.value2 ?? a.value), operator: op }
            : { formula1: String(a.value), operator: op };
          if (a.fill_color) cf.cellValue.format.fill.color = a.fill_color;
          if (a.font_color) cf.cellValue.format.font.color = a.font_color;
          break;
        }
        case "duplicate": {
          r.conditionalFormats.add(Excel.ConditionalFormatType.containsValues);
          // Альтернатива: presetCriteria с DuplicateValues
          break;
        }
        case "top10": {
          const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
          cf.topBottom.rule = { rank: a.value || 10, type: "TopItems" };
          if (a.fill_color) cf.topBottom.format.fill.color = a.fill_color;
          break;
        }
      }
      return;
    }

    case "create_chart": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const dataRange = sheet.getRange(a.data_range);
      const chart = sheet.charts.add(
        a.chart_type || "ColumnClustered",
        dataRange,
        a.series_by || "Auto"
      );
      if (a.title) { chart.title.text = a.title; chart.title.visible = true; }
      if (a.x_axis_title) { chart.axes.categoryAxis.title.text = a.x_axis_title; chart.axes.categoryAxis.title.visible = true; }
      if (a.y_axis_title) { chart.axes.valueAxis.title.text = a.y_axis_title; chart.axes.valueAxis.title.visible = true; }
      if (a.position_cell) {
        const target = sheet.getRange(a.position_cell);
        target.load("left,top");
        await context.sync();
        chart.left = target.left;
        chart.top = target.top;
      }
      if (a.width)  chart.width  = a.width;
      if (a.height) chart.height = a.height;
      return;
    }

    case "sort": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const r = sheet.getRange(a.range);
      r.sort.apply(
        [{ key: (a.key_column || 1) - 1, ascending: a.ascending !== false }],
        false, !!a.has_header
      );
      return;
    }

    case "insert_columns": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const col = a.before_column;
      const count = a.count || 1;
      const range = sheet.getRange(`${col}1:${col}1`).getEntireColumn();
      for (let i = 0; i < count; i++) range.insert("Right" === "Right" ? "Right" : "Left");
      return;
    }

    case "insert_rows": {
      const sheet = await resolveSheet(context, a, createdSheets);
      for (let i = 0; i < (a.count || 1); i++) {
        sheet.getRange(`${a.before_row}:${a.before_row}`).insert("Down");
      }
      return;
    }

    case "delete_columns": {
      const sheet = await resolveSheet(context, a, createdSheets);
      sheet.getRange(a.range).getEntireColumn().delete("Left"); return;
    }

    case "delete_rows": {
      const sheet = await resolveSheet(context, a, createdSheets);
      sheet.getRange(a.range).getEntireRow().delete("Up"); return;
    }

    case "data_validation": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const r = sheet.getRange(a.range);
      const v = r.dataValidation;
      if (a.rule === "list") {
        v.rule = { list: { inCellDropDown: true, source: (a.values || []).join(",") } };
      } else if (a.rule === "whole_number" || a.rule === "decimal") {
        const ruleSet = { wholeNumber: { formula1: a.min ?? 0, formula2: a.max ?? 100, operator: "Between" } };
        v.rule = a.rule === "decimal" ? { decimal: ruleSet.wholeNumber } : ruleSet;
      } else if (a.rule === "date") {
        v.rule = { date: { formula1: a.min || "1900-01-01", formula2: a.max || "2100-12-31", operator: "Between" } };
      }
      if (a.error_message)  v.errorAlert  = { message: a.error_message, showAlert: true, style: "Stop", title: "Ошибка ввода" };
      if (a.prompt_message) v.prompt      = { message: a.prompt_message, showPrompt: true, title: "Подсказка" };
      return;
    }

    case "protect_sheet": {
      const sheet = await resolveSheet(context, a, createdSheets);
      sheet.protection.protect({}, a.password || ""); return;
    }

    case "comment": {
      try {
        context.workbook.comments.add(
          `${(a.sheet === "_active_" ? context.workbook.worksheets.getActiveWorksheet().name : a.sheet)}!${a.cell}`,
          a.text
        );
      } catch (e) { console.warn("comment fail", e); }
      return;
    }

    case "hyperlink": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const cell = sheet.getRange(a.cell);
      cell.hyperlink = { address: a.url, textToDisplay: a.display || a.url };
      return;
    }

    case "clear": {
      const sheet = await resolveSheet(context, a, createdSheets);
      const r = sheet.getRange(a.range);
      if (a.contents && a.formats) r.clear("All");
      else if (a.formats) r.clear("Formats");
      else r.clear("Contents");
      return;
    }

    case "row_height": {
      const sheet = await resolveSheet(context, a, createdSheets);
      sheet.getRange(`${a.row}:${a.row}`).format.rowHeight = a.height; return;
    }

    case "column_width": {
      const sheet = await resolveSheet(context, a, createdSheets);
      sheet.getRange(`${a.column}:${a.column}`).format.columnWidth = a.width; return;
    }

    default:
      console.warn("Неизвестный тип действия:", a.type);
  }
}

/* ---------- ОСНОВНОЙ ЗАПУСК ---------- */

async function runAI() {
  const promptInput = $("promptInput");
  const licenseInput = $("licenseKey");
  const pin = licenseInput.value.trim();
  const prompt = promptInput.value.trim();

  if (!pin)    { setStatus("status-error", "❌ Введите лицензионный ключ."); return; }
  if (!prompt) { setStatus("status-error", "❌ Опишите задачу."); return; }

  setStatus("status-loading", "⏳ Считываю данные с листа...");
  $("runBtn").disabled = true;

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const used = sheet.getUsedRange();
      used.load("values");
      await context.sync();

      const values = used && used.values ? used.values : [];

      setStatus("status-loading", "🤖 ИИ обрабатывает запрос...");

      const response = await fetch(API_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          pin,
          prompt,
          data: values,
          sheet_name: sheet.name,
          mode: CURRENT_MODE,
        }),
      });

      const reply = await response.json();

      if (reply.status !== "success") {
        setStatus("status-error", "⚠️ " + (reply.message || "Неизвестная ошибка."));
        saveHistory({ prompt, ok: false, ts: Date.now() });
        return;
      }

      setStatus("status-loading", `🛠 Применяю ${reply.actions.length} действий...`);

      await executePlan(context, reply.actions);
      await context.sync();

      const meta = reply.meta || {};
      setStatus(
        "status-success",
        `✅ ${reply.message}\n` +
        `Действий: ${meta.actions_count ?? reply.actions.length}, ` +
        `попыток: ${meta.attempts ?? "—"}, ` +
        `время: ${meta.elapsed_sec ?? "—"} с`
      );
      saveHistory({ prompt, ok: true, ts: Date.now() });
    });
  } catch (e) {
    console.error(e);
    setStatus("status-error", "❌ Ошибка: " + e.message);
    saveHistory({ prompt, ok: false, ts: Date.now() });
  } finally {
    $("runBtn").disabled = false;
  }
}

/* ---------- ИНИЦИАЛИЗАЦИЯ ---------- */

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  // Восстанавливаем сохранённый ключ
  const savedPin = localStorage.getItem(STORAGE_KEYS.LICENSE);
  if (savedPin) $("licenseKey").value = savedPin;

  // Биндинги
  $("runBtn").onclick = runAI;
  const verifyBtn = $("verifyBtn");
  if (verifyBtn) verifyBtn.onclick = verifyLicense;

  const clearBtn = $("clearHistoryBtn");
  if (clearBtn) clearBtn.onclick = () => {
    localStorage.removeItem(STORAGE_KEYS.HISTORY);
    renderHistory();
  };

  // Ctrl+Enter для запуска
  $("promptInput").addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter") runAI();
  });

  renderHistory();
});
