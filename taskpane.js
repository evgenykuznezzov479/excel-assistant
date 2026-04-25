Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const prompt = document.getElementById("promptInput").value;
    const apiKey = document.getElementById("apiKey").value;
    const resultDiv = document.getElementById("result");

    if (!apiKey) { resultDiv.innerText = "Ошибка: Введите API ключ AI TUNNEL."; return; }
    
    resultDiv.innerText = "Анализирую данные...";

    try {
        await Excel.run(async (context) => {
            // 1. Собираем контекст: читаем данные выделенного диапазона
            const range = context.workbook.getSelectedRange();
            range.load("address, values");
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // 2. Системный промпт "Senior Data Analyst"
            const systemInstruction = `Ты эксперт по анализу данных в Excel (Office.js).
            Твоя задача: 
            1. Если просят аналитику: прочитай данные из 'range.values', проведи расчеты (JS), создай новый лист и запиши туда сводную таблицу.
            2. Если просят форматирование: используй 'range.format.fill.color'.
            3. Если просят функции: используй 'range.formulas = [["=SUM(...)"]]'.

            Контекст:
            - Выделенные данные: ${JSON.stringify(range.values)}
            - Адрес диапазона: ${range.address}
            - Существующие листы: ${sheets.items.map(s => s.name).join(", ")}

            Правила:
            - Всегда возвращай СТРОГО JSON: {"type": "code", "script": "ВАШ_КОД"}
            - Для создания листа используй: const s = context.workbook.worksheets.add("Имя"); s.activate();
            - НЕ пиши пояснительного текста, только JSON.
            
            Задача: ${prompt}`;

            // 3. Запрос к AI TUNNEL
            const response = await fetch("https://api.aitunnel.ru/v1/chat/completions", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
                body: JSON.stringify({ 
                    model: "gemini-2.5-flash",
                    messages: [{ role: "user", content: systemInstruction }] 
                })
            });

            const data = await response.json();
            const aiText = data.choices[0].message.content.replace(/```json/gi, "").replace(/```/g, "").trim();
            const aiResponse = JSON.parse(aiText);

            // 4. Выполнение
            if (aiResponse.type === "code") {
                const executeCode = new Function("context", `return (async () => { ${aiResponse.script} await context.sync(); })();`);
                await executeCode(context);
                resultDiv.innerText = "✅ Анализ завершен!";
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка: " + error.message;
    }
}
