Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const promptInput = document.getElementById("promptInput");
    const licenseInput = document.getElementById("licenseKey");
    const resultDiv = document.getElementById("result");

    resultDiv.className = "status-loading";
    resultDiv.innerText = "⏳ ИИ планирует задачи...";

    try {
        await Excel.run(async (context) => {
            const workbook = context.workbook;
            let sheet = workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            range.load("values");
            await context.sync();

            const response = await fetch("https://excel-ai-pro.ru/api/analyze", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    pin: licenseInput.value,
                    prompt: promptInput.value,
                    data: range.values
                })
            });

            const reply = await response.json();
            if (reply.status !== "success") throw new Error(reply.message);

            // 1. ПРИМЕНЯЕМ ОФОРМЛЕНИЕ (ACTIONS)
            for (const action of reply.actions) {
                if (action.type === "add_sheet") {
                    sheet = workbook.worksheets.add(action.name);
                }
                if (action.type === "rename") {
                    sheet.name = action.new_name;
                }
                if (action.type === "format") {
                    let target = sheet.getRange(action.range);
                    if (action.bold) target.format.font.bold = true;
                    if (action.bg) target.format.fill.color = action.bg;
                    if (action.color) target.format.font.color = action.color;
                }
                if (action.type === "chart") {
                    let source = sheet.getRange(action.source);
                    sheet.charts.add(action.chart_type, source, "Auto").title.text = action.title;
                }
            }

            // 2. ВСТАВЛЯЕМ ДАННЫЕ
            if (reply.new_data && reply.new_data.length > 0) {
                const targetRange = sheet.getRange("A1").getResizedRange(
                    reply.new_data.length - 1, 
                    reply.new_data[0].length - 1
                );
                targetRange.values = reply.new_data;
                targetRange.format.autofitColumns();
            }

            await context.sync();
            resultDiv.className = "status-success";
            resultDiv.innerText = "✅ " + reply.message;
        });
    } catch (e) {
        resultDiv.className = "status-error";
        resultDiv.innerText = "❌ Ошибка: " + e.message;
    }
}
