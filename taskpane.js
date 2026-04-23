Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const prompt = document.getElementById("promptInput").value;
    const apiKey = document.getElementById("apiKey").value;
    const resultDiv = document.getElementById("result");

    if (!apiKey) { resultDiv.innerText = "Пожалуйста, введите API ключ."; return; }
    if (!prompt) { resultDiv.innerText = "Пожалуйста, напишите задачу."; return; }
    
    resultDiv.innerText = "Думаю и пишу код...";

    try {
        await Excel.run(async (context) => {
            // 1. Собираем базовый контекст для ИИ
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

            // 2. Системный промпт (Режим программиста)
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
            3. Если пользователь просит аналитику с других листов: сначала прочитай данные (load -> sync), сделай расчеты средствами JS, затем запиши результат на новый или текущий лист.
            
            Формат ответа СТРОГО JSON:
            {"type": "code", "script": "const sheet = context.workbook.worksheets.getActiveWorksheet(); sheet.getRange('A1').values = [['Привет']];"}
            ИЛИ, если запрос не касается действий в Excel (просто вопрос):
            {"type": "message", "text": "Твой текстовый ответ"}`;

            // 3. Отправляем запрос к Gemini (Ваш оригинальный URL)
            const response = await fetch(`https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=${apiKey}`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ 
                    contents: [{ parts: [{ text: systemInstruction }] }],
                    // Добавленное улучшение: заставляем API вернуть чистый JSON
                    generationConfig: { 
                        responseMimeType: "application/json" 
                    }
                })
            });

            const data = await response.json();
            if (data.error) throw new Error(`API: ${data.error.message}`);
            if (!data.candidates || data.candidates.length === 0) throw new Error("Пустой ответ от сервера.");

            const aiText = data.candidates[0].content.parts[0].text;
            
            // 4. Исполнение ответа
            try {
                // Улучшение: парсим напрямую, так как API гарантирует валидный JSON без маркдауна
                const aiResponse = JSON.parse(aiText);

                if (aiResponse.type === "message") {
                    resultDiv.innerText = aiResponse.text;
                } else if (aiResponse.type === "code") {
                    resultDiv.innerText = "Выполняю скрипт в Excel...";
                    
                    // МАГИЯ: Динамическое выполнение кода, написанного нейросетью
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
                // Если ИИ ошибся в синтаксисе JSON или кода
                resultDiv.innerText = "Сбой при разборе ответа: " + e.message + "\n\nОтвет ИИ был: " + aiText;
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка: " + error.message;
    }
}
