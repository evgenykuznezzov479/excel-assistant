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
    if (!prompt) { resultDiv.innerText = "Ошибка: Напишите задачу."; return; }
    
    resultDiv.innerText = "Отправляю запрос...";

    try {
        // 1. Формируем системный промпт (как у вас было)
        // ВАЖНО: Мы не вызываем Excel.run СРАЗУ, чтобы сначала дождаться ответа от ИИ
        
        const systemInstruction = `Ты Senior Разработчик Office.js. Верни СТРОГО JSON: {"type": "code", "script": "..."} или {"type": "message", "text": "..."}. Запрос: ${prompt}`;

        // 2. Запрос к AI TUNNEL
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

        if (!response.ok) {
            throw new Error(`Ошибка сервера: ${response.status} ${response.statusText}`);
        }

        const data = await response.json();
        const aiText = data.choices[0].message.content;
        resultDiv.innerText = "Получил код от ИИ, выполняю...";

        // 3. Выполнение в Excel
        await Excel.run(async (context) => {
            const cleanJson = aiText.replace(/```json/gi, "").replace(/```javascript/gi, "").replace(/```/g, "").trim();
            const aiResponse = JSON.parse(cleanJson);

            if (aiResponse.type === "message") {
                resultDiv.innerText = aiResponse.text;
            } else {
                const executeCode = new Function("context", `return (async () => { ${aiResponse.script} })();`);
                await executeCode(context);
                await context.sync();
                resultDiv.innerText = "✅ Готово!";
            }
        });

    } catch (error) {
        console.error(error);
        resultDiv.innerText = "❌ Ошибка: " + error.message;
    }
}
