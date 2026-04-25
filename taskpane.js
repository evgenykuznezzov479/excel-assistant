Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const prompt = document.getElementById("promptInput").value;
    const apiKey = document.getElementById("apiKey").value;
    const resultDiv = document.getElementById("result");

    if (!apiKey) { resultDiv.innerText = "❌ Введите API ключ."; return; }
    resultDiv.innerText = "⏳ Анализирую таблицу...";

    try {
        await Excel.run(async (context) => {
            // 1. Берем используемый диапазон (автоматически находит всю таблицу)
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            range.load("values, address");
            await context.sync();

            const data = range.values;

            // 2. Системная инструкция для ИИ
            const systemInstruction = `Ты — эксперт Office.js. Твоя задача — написать JS-код для решения задачи.
            ЗАДАЧА: "${prompt}"
            ДАННЫЕ ТАБЛИЦЫ (первые 10 строк): ${JSON.stringify(data.slice(0, 10))}
            
            ПРАВИЛА ДЛЯ КОДА:
            - Используй ПЕРЕДАННЫЙ массив 'data' (это все данные таблицы).
            - Для поиска столбцов (Цена, Номенклатура и т.д.) сначала найди строку с заголовками в массиве 'data'.
            - ВМЕСТО range.rowCount используй data.length.
            - Для создания листа: const newSheet = context.workbook.worksheets.add("Результат_" + Date.now().toString().slice(-4));
            - Для записи данных: newSheet.getRange("A1:C10").values = [массив]; (указывай размер правильно).
            - ОБЯЗАТЕЛЬНО в конце: await context.sync();
            
            Верни СТРОГО JSON: {"type": "code", "script": "..."}`;

            const response = await fetch("https://api.aitunnel.ru/v1/chat/completions", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
                body: JSON.stringify({ 
                    model: "gemini-2.5-flash",
                    messages: [{ role: "user", content: systemInstruction }] 
                })
            });

            const aiData = await response.json();
            const aiResponse = JSON.parse(aiData.choices[0].message.content.replace(/```json|```|```javascript/gi, "").trim());

            if (aiResponse.type === "code") {
                resultDiv.innerText = "🚀 Выполняю...";
                
                // Передаем данные напрямую в функцию, чтобы избежать проблем с .load()
                const executeCode = new Function("context", "data", `
                    return (async () => {
                        try {
                            ${aiResponse.script}
                            await context.sync();
                        } catch (e) {
                            throw new Error("Ошибка в коде Excel: " + e.message);
                        }
                    })();
                `);

                await executeCode(context, data);
                resultDiv.innerText = "✅ Готово!";
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка: " + error.message;
        console.error(error);
    }
}
