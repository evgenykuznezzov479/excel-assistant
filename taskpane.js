async function runAI() {
    const prompt = document.getElementById("promptInput").value;
    const apiKey = document.getElementById("apiKey").value;
    const resultDiv = document.getElementById("result");

    if (!apiKey) { resultDiv.innerText = "Пожалуйста, введите API ключ."; return; }
    if (!prompt) { resultDiv.innerText = "Пожалуйста, напишите задачу."; return; }
    
    resultDiv.innerText = "Думаю и пишу код...";

    try {
        await Excel.run(async (context) => {
            // 1. Собираем базовый контекст
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            activeSheet.load("name");
            await context.sync();
            const sheetNames = sheets.items.map(s => s.name).join(", ");

            let selectedAddress = "Ничего не выделено";
            try {
                const range = context.workbook.getSelectedRange();
                range.load("address");
                await context.sync();
                selectedAddress = range.address;
            } catch (e) {}

            // 2. Системный промпт
            const systemInstruction = `Ты Senior Разработчик Office.js (Excel JavaScript API). 
            Твоя задача — переводить запросы пользователя в готовый, рабочий код на JavaScript, который будет выполнен внутри 'Excel.run(async (context) => { ... })'.
            
            Контекст книги пользователя:
            - Существующие листы: ${sheetNames}
            - Текущий активный лист: ${activeSheet.name}
            - Выделенный диапазон: ${selectedAddress}
            
            Запрос пользователя: "${prompt}"

            ПРАВИЛА ГЕНЕРАЦИИ КОДА:
            1. Используй только объект 'context'.
            2. Обязательно вызывай 'await context.sync()' после команд чтения (load) и перед чтением свойств.
            3. Если пользователь просит аналитику с других листов: сначала прочитай данные (load -> sync), сделай расчеты средствами JS, затем запиши результат.
            4. Верни СТРОГО формат JSON. Никакого маркдауна или пояснительного текста.

            Формат ответа JSON:
            {"type": "code", "script": "const sheet = context.workbook.worksheets.getActiveWorksheet(); sheet.getRange('A1').values = [['Привет']];"}
            ИЛИ
            {"type": "message", "text": "Твой текстовый ответ"}`;

            // 3. Отправляем запрос к AI TUNNEL (OpenAI-compatible)
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
            
            // Обработка ошибок API
            if (data.error) throw new Error(`API Error: ${data.error.message}`);
            if (!data.choices || data.choices.length === 0) throw new Error("Пустой ответ от сервера.");

            // Извлечение текста из ответа в формате OpenAI
            const aiText = data.choices[0].message.content;
            
            // 4. Исполнение ответа
            try {
                const cleanJson = aiText.replace(/```json/gi, "").replace(/```javascript/gi, "").replace(/```/g, "").trim();
                const aiResponse = JSON.parse(cleanJson);

                if (aiResponse.type === "message") {
                    resultDiv.innerText = aiResponse.text;
                } else if (aiResponse.type === "code") {
                    resultDiv.innerText = "Выполняю скрипт в Excel...";
                    
                    const executeCode = new Function("context", `
                        return (async () => {
                            try {
                                ${aiResponse.script}
                            } catch (err) {
                                throw new Error("Ошибка в сгенерированном коде: " + err.message);
                            }
                        })();
                    `);
                    
                    await executeCode(context);
                    await context.sync();
                    resultDiv.innerText = "✅ Готово!";
                }
            } catch (e) {
                resultDiv.innerText = "Сбой при разборе ответа: " + e.message + "\n\nОтвет ИИ был: " + aiText;
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка: " + error.message;
    }
}
