Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Привязываем функцию к кнопке в интерфейсе
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const promptInput = document.getElementById("promptInput");
    const licenseInput = document.getElementById("licenseKey");
    const resultDiv = document.getElementById("result");

    // Визуальная индикация загрузки
    resultDiv.className = "status-loading";
    resultDiv.innerText = "⏳ ИИ обрабатывает данные...";

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            
            // Загружаем значения ячеек
            range.load("values");
            await context.sync();

            const tableData = range.values;

            // Отправка запроса на ваш сервер
            const response = await fetch("https://excel-ai-pro.ru/api/analyze", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    pin: licenseInput.value,
                    prompt: promptInput.value,
                    data: tableData
                })
            });

            const reply = await response.json();

            if (reply.status === "success") {
                // Создаем новый лист для вывода результата
                const newSheet = context.workbook.worksheets.add("Результат ИИ");
                
                // Определяем диапазон для вставки данных
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
                resultDiv.innerText = "⚠️ Ошибка: " + reply.message;
            }
            
            await context.sync();
        });
    } catch (error) {
        resultDiv.className = "status-error";
        resultDiv.innerText = "❌ Критическая ошибка: " + error.message;
    }
}
