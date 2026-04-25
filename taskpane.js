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
    resultDiv.innerText = "Анализирую и выполняю...";

    try {
        await Excel.run(async (context) => {
            // 1. Получаем выделенный диапазон
            const range = context.workbook.getSelectedRange();
            range.load("address, values");
            await context.sync();

            const data = range.values;

            // 2. Системный промпт (Универсальный)
            const systemInstruction = `Ты — Senior Developer надстроек Office.js.
            Задача пользователя: "${prompt}"
            Текущие данные (первые 5 строк для структуры): ${JSON.stringify(data.slice(0, 5))}
            
            Твоя задача — вернуть СТРОГО JSON: {"type": "code", "script": "ТВОЙ_JS_КОД"}
            
            ИНСТРУКЦИИ ДЛЯ КОДА:
            - Если нужно создать лист: context.workbook.worksheets.add("Имя").
            - Если нужно фильтровать 5000+ строк: используй JS-фильтрацию (data.filter), а не формулы Excel.
            - Если нужно записать данные: sheet.getRange("A1").values = [массив_данных].
            - Если нужно форматирование: range.format.fill.color = "#ЦВЕТ".
            - ВАЖНО: Весь код должен быть внутри await Excel.run, но НЕ пиши 'await Excel.run' внутри скрипта.
            - Обязательно в конце кода: await context.sync();
            - Никаких markdown-тегов (```json), только чистый JSON.`;

            // 3. Запрос к AI TUNNEL
            const response = await fetch("[https://api.aitunnel.ru/v1/chat/completions](https://api.aitunnel.ru/v1/chat/completions)", {
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
                // Выполняем код, присланный ИИ
                const executeCode = new Function("context", "data", `
                    return (async () => {
                        try {
                            ${aiResponse.script}
                        } catch (err) {
                            throw new Error("Ошибка в скрипте: " + err.message);
                        }
                    })();
                `);
                
                await executeCode(context, data);
                resultDiv.innerText = "✅ Выполнено!";
            } else {
                resultDiv.innerText = aiResponse.text || "Готово.";
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка: " + error.message;
        console.error("DEBUG:", error);
    }
}
