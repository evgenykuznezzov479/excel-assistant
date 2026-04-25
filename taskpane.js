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
    resultDiv.innerText = "Анализирую таблицу...";

    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("address, values");
            await context.sync();

            const data = range.values;
            // Передаем в ИИ всю таблицу, чтобы он сам понял структуру
            const tableSummary = JSON.stringify(data.slice(0, 10)); 

            const systemInstruction = `Ты — универсальный инженер данных Excel.
            Задача пользователя: "${prompt}"
            
            СТРУКТУРА ТАБЛИЦЫ (Заголовки и данные):
            ${tableSummary}
            
            ТВОЙ АЛГОРИТМ:
            1. Проанализируй заголовки (строка 0).
            2. Найди индексы столбцов, которые семантически соответствуют запросу пользователя. 
               - Например: "Цена" может называться "Стоимость", "Price", "Amount".
               - "Товар" может быть "Номенклатура", "Название", "Item".
            3. Если не уверен в заголовке — проанализируй типы данных в колонках (где числа, где текст).
            4. Сгенерируй JavaScript-код, который использует динамические переменные для индексов найденных столбцов.
            
            Верни СТРОГО JSON: {"type": "code", "script": "ТВОЙ_JS_КОД"}
            - Используй 'data' для работы.
            - Обязательно в конце: await context.sync();
            - НЕ пиши пояснений.`;

            const response = await fetch("https://api.aitunnel.ru/v1/chat/completions", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
                body: JSON.stringify({ 
                    model: "gemini-2.5-flash",
                    messages: [{ role: "user", content: systemInstruction }] 
                })
            });

            const aiData = await response.json();
            const aiText = aiData.choices[0].message.content.replace(/```json|```|```javascript/gi, "").trim();
            const aiResponse = JSON.parse(aiText);

            if (aiResponse.type === "code") {
                const executeCode = new Function("context", "data", `return (async () => { ${aiResponse.script} })();`);
                await executeCode(context, data);
                resultDiv.innerText = "✅ Выполнено!";
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка: " + error.message;
    }
}
