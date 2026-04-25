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
    if (!prompt) { resultDiv.innerText = "Ошибка: Напишите задачу."; return; }
    
    resultDiv.innerText = "Выполняю задачу...";

    try {
        await Excel.run(async (context) => {
            // 1. Получаем данные выделения
            const range = context.workbook.getSelectedRange();
            range.load("address, values");
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // 2. Системный промпт
            const systemInstruction = `Ты — Senior Developer надстроек Office.js.
            Задача пользователя: "${prompt}"
            Данные выделения: ${JSON.stringify(range.values)}
            
            Твоя задача — вернуть СТРОГО JSON: {"type": "code", "script": "ТВОЙ_JS_КОД"}
            
            ВАЖНЫЕ ПРАВИЛА:
            - Используй context.workbook.worksheets.add("Имя") для создания листа.
            - Используй range.format.fill.color для цвета.
            - Для наценки 50% и фильтрации: пиши JS-код, который берет массив 'data', делает filter/map и записывает результат через sheet.getRange().values = ...
            - Код должен заканчиваться действием (запись данных или форматирование).
            - Никаких пояснений, только JSON.`;

            // 3. Запрос к AI TUNNEL
            const response = await fetch("https://api.aitunnel.ru/v1/chat/completions", {
                method: "POST",
                headers: { 
                    "Content-Type": "application/json", 
                    "Authorization": `Bearer ${apiKey}` 
                },
                body: JSON.stringify({ 
                    model: "gemini-2.5-flash",
                    messages: [{ role: "user", content: systemInstruction }] 
                })
            });

            const data = await response.json();
            if (!data.choices || !data.choices[0]) throw new Error("Нет ответа от ИИ.");
            
            const aiText = data.choices[0].message.content.replace(/```json|```javascript|```/gi, "").trim();
            const aiResponse = JSON.parse(aiText);

            // 4. Исполнение кода
            if (aiResponse.type === "code") {
                // Создаем функцию, принудительно делающую sync в конце
                const executeCode = new Function("context", "data", `
                    return (async () => {
                        console.log("Начало выполнения скрипта ИИ");
                        ${aiResponse.script}
                        await context.sync();
                        console.log("Скрипт успешно завершен и синхронизирован");
                    })();
                `);

                await executeCode(context, range.values);
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
