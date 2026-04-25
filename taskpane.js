Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const prompt = document.getElementById("promptInput").value;
    const apiKey = document.getElementById("apiKey").value;
    const resultDiv = document.getElementById("result");

    if (!apiKey) { resultDiv.innerText = "Ошибка: Введите API ключ."; return; }
    resultDiv.innerText = "Выполняю задачу...";

    try {
        await Excel.run(async (context) => {
            // 1. Собираем контекст: читаем заголовки и данные
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = context.workbook.getSelectedRange();
            range.load("address, values");
            sheet.load("name");
            await context.sync();

            // 2. Системный промпт "Универсальный Инженер Excel"
            const systemInstruction = `Ты Senior Developer надстроек Office.js. 
            Твоя задача — выполнять любые действия в Excel по запросу пользователя.

            ТВОИ ВОЗМОЖНОСТИ:
            1. АНАЛИТИКА: Используй переданный массив данных (data), фильтруй его, считай наценки/суммы (JS), записывай результат на новый или существующий лист.
            2. ФОРМАТИРОВАНИЕ: Меняй цвета (range.format.fill.color), шрифты, границы.
            3. СТРУКТУРА: Создавай листы (context.workbook.worksheets.add), переименовывай их.
            4. ФОРМУЛЫ: Вставляй формулы Excel (range.formulas = [["=SUM(...)"]]).

            Контекст:
            - Адрес выделения: ${range.address}
            - Данные (первые 50 строк): ${JSON.stringify(range.values.slice(0, 50))}
            
            ПРАВИЛА:
            - Не пиши пояснений. Верни СТРОГО JSON: {"type": "code", "script": "ВАШ_КОД"}.
            - Всегда используй 'await context.sync()' в конце скрипта.
            - Для фильтрации 5000+ строк используй методы JS (filter, map, reduce).

            ЗАПРОС ПОЛЬЗОВАТЕЛЯ: ${prompt}`;

            // 3. Запрос к AI TUNNEL
            const response = await fetch("https://api.aitunnel.ru/v1/chat/completions", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
                body: JSON.stringify({ 
                    model: "gemini-2.5-flash",
                    messages: [{ role: "user", content: systemInstruction }] 
                })
            });

            const aiData = await response.json();
            const aiText = aiData.choices[0].message.content.replace(/```json|```javascript|```/gi, "").trim();
            const aiResponse = JSON.parse(aiText);

            // 4. Исполнение
            if (aiResponse.type === "code") {
                const executeCode = new Function("context", "data", `return (async () => { ${aiResponse.script} await context.sync(); })();`);
                await executeCode(context, range.values);
                resultDiv.innerText = "✅ Выполнено!";
            } else {
                resultDiv.innerText = aiResponse.text || "Готово.";
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка: " + error.message;
        console.error(error);
    }
}
