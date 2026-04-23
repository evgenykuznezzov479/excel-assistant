// ... (предыдущий код до Шага 4 остается без изменений) ...

            // 4. Запрос к Gemini
            const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: systemInstruction }] }]
                })
            });

            const data = await response.json();

            // === НОВАЯ ЛОГИКА ПРОВЕРКИ ОШИБОК ===
            // Если Google вернул ошибку вместо данных
            if (data.error) {
                throw new Error(`Google API: ${data.error.message}`);
            }
            // Если ответ пустой (например, сработал фильтр безопасности)
            if (!data.candidates || data.candidates.length === 0) {
                 throw new Error(`Google не вернул ответ. Возможно, сработал фильтр или проблема с сетью. Ответ сервера: ${JSON.stringify(data)}`);
            }
            // ===================================

            const aiText = data.candidates[0].content.parts[0].text;

            // 5. Исполнение команд
            try {
                // Пытаемся очистить ответ от маркдауна и распарсить JSON
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
                // Если ИИ ответил просто текстом, выводим его
                resultDiv.innerText = aiText;
            }
        });
    } catch (error) {
        // Теперь панель покажет вам настоящую причину ошибки
        resultDiv.innerText = "❌ Ошибка: " + error.message;
    }
}