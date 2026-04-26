Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const prompt = document.getElementById("promptInput").value;
    const resultDiv = document.getElementById("result");

    if (!prompt) { resultDiv.innerText = "❌ Напишите задачу!"; return; }
    resultDiv.innerText = "🚀 Анализ на сервере...";

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            range.load("values");
            await context.sync();

            const data = range.values;

            // Отправляем на сервер (ВАША ССЫЛКА NGROK ДОЛЖНА БЫТЬ ЗДЕСЬ)
            const response = await fetch("https://rebuilt-nutmeg-breeches.ngrok-free.dev/api/analyze", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ prompt: prompt, data: data })
            });

            const serverReply = await response.json();

            if (serverReply.status === "success") {
                resultDiv.innerText = "✅ Данные получены. Создаю лист...";
                
                // Вставляем данные, если ИИ их вернул
                if (serverReply.new_data && serverReply.new_data.length > 0) {
                    const rowCount = serverReply.new_data.length;
                    const colCount = serverReply.new_data[0].length;
                    
                    // Создаем новый лист с уникальным именем
                    const newSheetName = "Анализ_" + Math.floor(Math.random() * 1000);
                    const newSheet = context.workbook.worksheets.add(newSheetName);
                    
                    // Безопасно выделяем диапазон нужного размера и вставляем данные
                    const targetRange = newSheet.getRange("A1").getResizedRange(rowCount - 1, colCount - 1);
                    targetRange.values = serverReply.new_data;
                    targetRange.format.autofitColumns();
                    
                    newSheet.activate();
                    await context.sync();
                    
                    resultDiv.innerText = "✅ Готово! Результат на листе " + newSheetName;
                }
            } else {
                resultDiv.innerText = "⚠️ Ошибка: " + serverReply.message;
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка Excel: " + error.message;
        console.error(error);
    }
}
