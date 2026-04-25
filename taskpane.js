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
    resultDiv.innerText = "ИИ анализирует структуру таблицы...";

    try {
        await Excel.run(async (context) => {
            // 1. Читаем всё выделение вместе с заголовками
            const range = context.workbook.getSelectedRange();
            range.load("address, values");
            await context.sync();

            const allData = range.values;
            const headers = allData[0]; // Первая строка - это заголовки

            // 2. Системный промпт "Универсальный аналитик"
            const systemInstruction = `Ты — эксперт по автоматизации Excel.
            Задача пользователя: "${prompt}"
            
            СТРУКТУРА ТАБЛИЦЫ:
            Заголовки столбцов: ${JSON.stringify(headers)}
            Данные (первые 5 строк): ${JSON.stringify(allData.slice(1, 6))}
            
            Твои правила:
            1. ПЕРВЫМ ДЕЛОМ определи индексы столбцов для "Название" (или номенклатура) и "Цена" (или стоимость) по заголовкам.
            2. Если нужно создать отчет: создай новый лист (context.workbook.worksheets.add), и запиши туда результат.
            3. Если нужна аналитика (средние, суммы): делай расчеты в JS на массиве данных.
            4. Если нужна фильтрация: делай ее в JS, учитывая найденные индексы столбцов.
            5. Для наценки: умножай найденное значение цены на 1.5.
            
            Верни СТРОГО JSON: {"type": "code", "script": "ТВОЙ_JS_КОД"}
            В коде используй переменную 'data' (это весь диапазон). 
            ОБЯЗАТЕЛЬНО закончи код: await context.sync();`;

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
                resultDiv.innerText = "Применяю логику...";
                const executeCode = new Function("context", "data", `
                    return (async () => {
                        try {
                            ${aiResponse.script}
                        } catch (e) {
                            throw new Error("Ошибка в скрипте: " + e.message);
                        }
                    })();
                `);
                await executeCode(context, allData);
                resultDiv.innerText = "✅ Готово!";
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка: " + error.message;
        console.error("DEBUG:", error);
    }
}
