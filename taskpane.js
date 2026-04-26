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
    resultDiv.innerText = "⏳ Обработка...";

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
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

            // Если сервер выдаст ошибку Nginx (HTML), этот блок поможет её поймать
            if (!response.ok) {
                throw new Error("Сервер ответил ошибкой " + response.status);
            }

            const reply = await response.json();

            if (reply.status === "success") {
                const newSheet = context.workbook.worksheets.add("Результат ИИ");
                const targetRange = newSheet.getRange("A1").getResizedRange(
                    reply.new_data.length - 1, 
                    reply.new_data[0].length - 1
                );
                targetRange.values = reply.new_data;
                targetRange.format.autofitColumns();
                newSheet.activate();
                
                resultDiv.className = "status-success";
                resultDiv.innerText = "✅ " + reply.message;
            } else {
                resultDiv.className = "status-error";
                resultDiv.innerText = "⚠️ " + reply.message;
            }
            await context.sync();
        });
    } catch (error) {
        resultDiv.className = "status-error";
        resultDiv.innerText = "❌ Ошибка: " + error.message;
    }
}
