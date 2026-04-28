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
  MODE: "aiapro_mode",
};

let CURRENT_MODE = "auto";

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

/* ---------- РЕЖИМЫ ---------- */

function setMode(mode) {
  CURRENT_MODE = mode;
  document.querySelectorAll(".mode-chip").forEach((el) => {
    el.classList.toggle("active", el.dataset.mode === mode);
  });
  localStorage.setItem(STORAGE_KEYS.MODE, mode);
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
      setStatus("status-success",
        `✅ ${j.name} | Тариф: ${j.tier} | ` +
        (j.limit_per_day ? `Использовано сегодня: ${j.used_today}/${j.limit_per_day}` : "Без лимита"));
    } else {
      setStatus("status-error", "⚠️ " + (j.message || "Лицензия не найдена."));
    }
  } catch (e) {
    setStatus("status-error", "❌ Сервер недоступен: " + e.message);
  }
}

/* ---------- ИСПОЛНИТЕЛЬ ACTION PLAN ---------- */

function resolveSheet(context, action, createdSheets) {
  // sheet может быть: "_active_", "_new_", или "Имя листа"
  const wb = context.workbook;
  const name = action.sheet || "_active_";

  if (name === "_active_") {
    return wb.worksheets.getActiveWorksheet();
  }
  if (name === "_new_") {
    const desired = action.sheet_name || `Результат ИИ`;
    const unique = uniqueSheetName(context, desired);
    const s = wb.worksheets.add(unique);
    createdSheets.set(name, s);
    action.sheet = unique;
    return s;
  }
  // Существующий лист или новый по имени
  let s;
  try { s = wb.worksheets.getItem(name); }
  catch (_) { s = wb.worksheets.add(name); }
  return s;
}

function uniqueSheetName(context, base) {
  // Office.js не даёт синхронно проверить — добавим суффикс времени, если потребуется
  const stamp = new Date().toLocaleTimeString("ru-RU").replace(/[: ]/g, "");
  return `${base} ${stamp}`.slice(0, 30);
}

function colLetter(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s || "A";
}

async function executePlan(context, actions) {
  const createdSheets = new Map();
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
}

async function runAction(context, a, createdSheets) {
  switch (a.type) {

    case "create_sheet": {
      let sheet;
      try { sheet = context.workbook.worksheets.getItem(a.name); }
      catch { sheet = context.workbook.worksheets.add(a.name); }
      if (a.color) sheet.tabColor = a.color;
      if (a.activate !== false) sheet.activate();
      return;
    }

    case "rename_sheet": {
      const s = context.workbook.worksheets.getItem(a.old_name);
      s.name = a.new_name; return;
    }

    case "delete_sheet": {
      try { context.workbook.worksheets.getItem(a.name).delete(); } catch {}
      return;
    }

    case "activate_sheet": {
      context.workbook.worksheets.getItem(a.name).activate(); return;
    }

    case "set_values": {
      const sheet = resolveSheet(context, a, createdSheets);
      const values = a.values || [];
      if (!values.length) return;
      const rows = values.length;
      const cols = values[0].length;
      const startCell = a.start_cell || "A1";
      const range = sheet.getRange(startCell).getResizedRange(rows - 1, cols - 1);
      range.values = values;
      return;
    }

    case "set_formula": {
      const sheet = resolveSheet(context, a, createdSheets);
      sheet.getRange(a.cell).formulas = [[a.formula]];
      return;
    }

    case "fill_formula": {
      const sheet = resolveSheet(context, a, createdSheets);
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
      const sheet = resolveSheet(context, a, createdSheets);
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
      const sheet = resolveSheet(context, a, createdSheets);
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
      const sheet = resolveSheet(context, a, createdSheets);
      const ur = sheet.getUsedRange();
      if (a.mode === "rows") ur.format.autofitRows();
      else if (a.mode === "both") { ur.format.autofitColumns(); ur.format.autofitRows(); }
      else ur.format.autofitColumns();
      return;
    }

    case "merge_cells": {
      const sheet = resolveSheet(context, a, createdSheets);
      sheet.getRange(a.range).merge(!!a.across); return;
    }

    case "freeze_panes": {
      const sheet = resolveSheet(context, a, createdSheets);
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
      const sheet = resolveSheet(context, a, createdSheets);
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
      const sheet = resolveSheet(context, a, createdSheets);
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
      const sheet = resolveSheet(context, a, createdSheets);
      const r = sheet.getRange(a.range);
      r.sort.apply(
        [{ key: (a.key_column || 1) - 1, ascending: a.ascending !== false }],
        false, !!a.has_header
      );
      return;
    }

    case "insert_columns": {
      const sheet = resolveSheet(context, a, createdSheets);
      const col = a.before_column;
      const count = a.count || 1;
      const range = sheet.getRange(`${col}1:${col}1`).getEntireColumn();
      for (let i = 0; i < count; i++) range.insert("Right" === "Right" ? "Right" : "Left");
      return;
    }

    case "insert_rows": {
      const sheet = resolveSheet(context, a, createdSheets);
      for (let i = 0; i < (a.count || 1); i++) {
        sheet.getRange(`${a.before_row}:${a.before_row}`).insert("Down");
      }
      return;
    }

    case "delete_columns": {
      const sheet = resolveSheet(context, a, createdSheets);
      sheet.getRange(a.range).getEntireColumn().delete("Left"); return;
    }

    case "delete_rows": {
      const sheet = resolveSheet(context, a, createdSheets);
      sheet.getRange(a.range).getEntireRow().delete("Up"); return;
    }

    case "data_validation": {
      const sheet = resolveSheet(context, a, createdSheets);
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
      const sheet = resolveSheet(context, a, createdSheets);
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
      const sheet = resolveSheet(context, a, createdSheets);
      const cell = sheet.getRange(a.cell);
      cell.hyperlink = { address: a.url, textToDisplay: a.display || a.url };
      return;
    }

    case "clear": {
      const sheet = resolveSheet(context, a, createdSheets);
      const r = sheet.getRange(a.range);
      if (a.contents && a.formats) r.clear("All");
      else if (a.formats) r.clear("Formats");
      else r.clear("Contents");
      return;
    }

    case "row_height": {
      const sheet = resolveSheet(context, a, createdSheets);
      sheet.getRange(`${a.row}:${a.row}`).format.rowHeight = a.height; return;
    }

    case "column_width": {
      const sheet = resolveSheet(context, a, createdSheets);
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

/* ---------- ПРЕСЕТЫ БЫСТРЫХ ЗАДАЧ ---------- */

const PRESETS = [
  { icon: "📊", title: "Полный анализ",
    text: "Сделай полный анализ таблицы: выведи на новый лист сводку по ключевым метрикам, топ-5 строк по основным показателям, итоги по группам и построй 2-3 подходящие диаграммы.",
    mode: "analyze" },
  { icon: "🎨", title: "Зебра + заголовки",
    text: "Сделай красивое оформление таблицы: цветную шапку, чередующиеся строки (зебра), границы, автоширину столбцов, закрепи верхнюю строку.",
    mode: "format" },
  { icon: "🔥", title: "Тепловая карта",
    text: "Примени условное форматирование (color scale) ко всем числовым столбцам — от зелёного к красному.",
    mode: "format" },
  { icon: "Σ", title: "Добавить итоги",
    text: "Добавь строку 'Итого' внизу с формулами SUM по числовым столбцам и пометь её жирным.",
    mode: "formula" },
  { icon: "📈", title: "Диаграмма по столбцам",
    text: "Построй подходящую диаграмму на основе данных и помести её рядом с таблицей.",
    mode: "chart" },
  { icon: "🧹", title: "Очистка",
    text: "Удали пустые строки, обрежь лишние пробелы в ячейках, приведи заголовки к единому стилю Title Case.",
    mode: "transform" },
  { icon: "💲", title: "Наценка 15%",
    text: "Найди столбец с ценой и добавь рядом столбец 'Цена с наценкой 15%' с формулой.",
    mode: "formula" },
  { icon: "🔍", title: "Дубли",
    text: "Подсветь красным дубликаты в первом столбце и выведи на новом листе список уникальных значений с количеством повторов.",
    mode: "format" },
];

function renderPresets() {
  const wrap = $("presets");
  if (!wrap) return;
  wrap.innerHTML = "";
  PRESETS.forEach((p) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "preset";
    b.innerHTML = `<span>${p.icon}</span> ${p.title}`;
    b.onclick = () => {
      $("promptInput").value = p.text;
      setMode(p.mode);
      $("promptInput").focus();
    };
    wrap.appendChild(b);
  });
}

/* ---------- ИНИЦИАЛИЗАЦИЯ ---------- */

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  // Восстанавливаем сохранённые значения
  const savedPin = localStorage.getItem(STORAGE_KEYS.LICENSE);
  if (savedPin) $("licenseKey").value = savedPin;

  const savedMode = localStorage.getItem(STORAGE_KEYS.MODE) || "auto";
  setMode(savedMode);

  // Биндинги
  $("runBtn").onclick = runAI;
  const verifyBtn = $("verifyBtn"); if (verifyBtn) verifyBtn.onclick = verifyLicense;

  document.querySelectorAll(".mode-chip").forEach((el) => {
    el.onclick = () => setMode(el.dataset.mode);
  });

  const clearBtn = $("clearHistoryBtn");
  if (clearBtn) clearBtn.onclick = () => {
    localStorage.removeItem(STORAGE_KEYS.HISTORY);
    renderHistory();
  };

  // Ctrl+Enter для запуска
  $("promptInput").addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter") runAI();
  });

  renderPresets();
  renderHistory();
});
