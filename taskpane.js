Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const prompt = document.getElementById("promptInput").value;
    const resultDiv = document.getElementById("result");

    if (!prompt) { resultDiv.innerText = "❌ Напишите задачу!"; return; }
    resultDiv.innerText = "🚀 Отправляю данные на сервер...";

    try {
        await Excel.run(async (context) => {
            // 1. Берем всю таблицу с текущего листа
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            range.load("values");
            await context.sync();

            const data = range.values;

            // 2. Отправляем данные на ВАШ Python-сервер через ngrok
            // ВАЖНО: Убедитесь, что /api/analyze есть в конце ссылки!
            const response = await fetch("https://rebuilt-nutmeg-breeches.ngrok-free.dev/api/analyze", {
                method: "POST",
                headers: { 
                    "Content-Type": "application/json" 
                },
                body: JSON.stringify({ 
                    prompt: prompt, 
                    data: data 
                })
            });

            // 3. Получаем ответ от сервера
            const serverReply = await response.json();

            // 4. Выводим результат в панель
            if (serverReply.status === "success") {
                resultDiv.innerText = "✅ Ответ сервера: " + serverReply.message;
            } else {
                resultDiv.innerText = "⚠️ Ошибка сервера: " + JSON.stringify(serverReply);
            }
        });
    } catch (error) {
        resultDiv.innerText = "❌ Ошибка Excel: " + error.message;
        console.error(error);
    }
}
