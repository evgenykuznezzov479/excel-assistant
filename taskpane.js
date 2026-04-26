Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const prompt = document.getElementById("promptInput").value;
    const licenseKey = document.getElementById("licenseKey").value;
    const resultDiv = document.getElementById("result");

    resultDiv.className = "status-loading";
    resultDiv.innerText = "⏳ ИИ анализирует всю книгу...";

    try {
        await Excel.run(async (context) => {
            // 1. СОБИРАЕМ ДАННЫЕ СО ВСЕХ ЛИСТОВ
            let sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            let workbookData = {};
            for (let sheet of sheets.items) {
                let usedRange = sheet.getUsedRange();
                usedRange.load("values");
                await context.sync();
                workbookData[sheet.name] = usedRange.values;
            }

            // 2. ОТПРАВЛЯЕМ НА СЕРВЕР
            const response = await fetch("https://excel-ai-pro.ru/api/analyze", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    pin: licenseKey,
                    prompt: prompt,
                    workbook_data: workbookData
                })
            });

            const reply = await response.json();
            if (reply.status !== "success") throw new Error(reply.message);

            // 3. ВЫПОЛНЯЕМ КОМАНДЫ (ACTIONS)
            resultDiv.innerText = "🚀 Выполнение команд...";
            const plan = reply.plan;

            for (const action of plan.actions) {
                let targetSheet;
                if (action.type === "add_sheet") {
                    targetSheet = context.workbook.worksheets.add(action.name);
                } else {
                    targetSheet = context.workbook.worksheets.getItem(action.sheet);
                }

                if (action.type === "set_values") {
                    let range = targetSheet.getRange(action.range);
                    range.values = action.values;
                }

                if (action.type === "format") {
                    let range = targetSheet.getRange(action.range);
                    if (action.bold) range.format.font.bold = true;
                    if (action.color) range.format.font.color = action.color;
                    if (action.bg) range.format.fill.color = action.bg;
                    if (action.number_format) range.numberFormat = [[action.number_format]];
                }

                if (action.type === "formula") {
                    let range = targetSheet.getRange(action.range);
                    range.formulasR1C1 = [[action.formula]];
                }

                if (action.type === "chart") {
                    let source = targetSheet.getRange(action.source);
                    let chart = targetSheet.charts.add(action.chart_type, source, "Auto");
                    chart.title.text = action.title;
                }

                if (action.type === "autofit") {
                    targetSheet.getRange(action.range).format.autofitColumns();
                }
            }

            await context.sync();
            resultDiv.className = "status-success";
            resultDiv.innerText = "✅ " + plan.message;
        });
    } catch (e) {
        resultDiv.className = "status-error";
        resultDiv.innerText = "❌ Ошибка: " + e.message;
    }
}
