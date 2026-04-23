Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const prompt = document.getElementById("promptInput").value;
    const apiKey = document.getElementById("apiKey").value;
    const resultDiv = document.getElementById("result");

    if (!apiKey) { 
        resultDiv.innerText = "Пожалуйста, введите API ключ."; 
        return; 
    }
    
    // Если этот текст появится, значит кнопка работает!
    resultDiv.innerText = "Сканирую структуру книги и отправляю данные...";

    try {
        await Excel.run(async (context) => {
            // 1. Собираем имена листов
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();
            const sheetNames = sheets.items.map(s => s.name).join(", ");

            // 2. Пытаемся взять выделенные данные
            let selectedData = "Данные не выделены";
            try {
                const range = context.workbook.getSelectedRange();
                range.load("values");
                await context.sync();
                selectedData = JSON.stringify(range.values);
            } catch (e) {
                // Игнорируем ошибку, если ничего не выделено
            }

            // 3. Формируем инструкцию
            const systemInstruction = `Ты умный ассистент для Excel. Текущие листы в книге: ${sheetNames}. 
            Выделенные данные: ${selectedData}. 
            Задача пользователя: ${prompt}. 
            Если нужно выполнить действия (создать лист, записать данные), верни СТРОГО формат JSON без лишнего текста:
            {"actions": [{"type": "addSheet", "name": "Имя_Листа"}, {"type": "writeValue", "sheet": "Имя_Листа", "address": "A1", "value": "Текст или число"}]}
            Если это просто вопрос, ответь текстом.`;

            // 4. Отправляем запрос (VPN должен быть включен в режиме TUN)
            const response = await fetch(`https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=${apiKey}`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: systemInstruction }] }]
                })
            });

            const data = await response.json();

            // Проверки на ошибки от Google
            if (data.error) {
                throw new Error(`Ответ API: ${data.error.message}`);
            }
            if (!data.candidates || data.candidates.length === 0) {
                 throw new Error(`Пустой ответ от сервера. Возможно, сработал фильтр. Ответ: ${JSON.stringify(data)}`);
            }

            const aiText = data.candidates[0].content.parts[0].text;

            // 5. Выполняем команды
            try {
                const cleanJson = aiText.replace(/```json/g, "").replace(/```/g, "").trim();
                const command = JSON.parse(cleanJson);

                if (command.actions) {
                    resultDiv.innerText = "Применяю изменения...";
                    for (let action of command.actions) {
                        if (action.type === "addSheet") {
                            const existingSheet = sheets.items.find(s => s.name === action.name);
                            if (!existingSheet) {
                                context.workbook.worksheets.add(action.name);
                            }
                        }
                        if (action.type === "writeValue") {
                            const targetSheet = context.workbook.worksheets.getItem(action.sheet);
                            targetSheet.getRange(action.address).values = [[action.value]];
                        }
                    }
                    await context.sync();
                    resultDiv.innerText = "✅ Задача успешно выполнена!";
                }
            } catch (e) {
                // Выводим просто текст, если ИИ не сгенерировал JSON
                resultDiv.innerText = aiText;
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка: " + error.message;
    }
}
