Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Загружаем ключ из памяти
        const savedKey = localStorage.getItem("ai_pro_key");
        if (savedKey) {
            document.getElementById("licenseKey").value = savedKey;
        }
        document.getElementById("runBtn").onclick = runAI;
    }
});

async function runAI() {
    const prompt = document.getElementById("promptInput").value;
    const licenseKey = document.getElementById("licenseKey").value;
    const resultDiv = document.getElementById("result");

    if (!licenseKey) { 
        resultDiv.className = "status-error";
        resultDiv.innerText = "❌ Ошибка: Введите лицензионный ключ."; 
        return; 
    }
    if (!prompt) { 
        resultDiv.className = "status-error";
        resultDiv.innerText = "❌ Ошибка: Напишите задачу."; 
        return; 
    }

    // Сохраняем ключ
    localStorage.setItem("ai_pro_key", licenseKey);
    
    resultDiv.className = "status-loading";
    resultDiv.innerText = "⏳ Сервер ИП Посаднев обрабатывает данные...";

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            range.load("values");
            await context.sync();

            const response = await fetch("https://excel-ai-pro.ru/api/analyze", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ 
                    pin: licenseKey,
                    prompt: prompt, 
                    data: range.values 
                })
            });

            const reply = await response.json();

            if (reply.status === "success") {
                resultDiv.className = "status-success";
                resultDiv.innerText = "✅ " + reply.message;

                if (reply.new_data && reply.new_data.length > 0) {
                    const newSheetName = "Результат_" + Math.floor(Math.random() * 900 + 100);
                    const newSheet = context.workbook.worksheets.add(newSheetName);
                    const target = newSheet.getRange("A1").getResizedRange(reply.new_data.length - 1, reply.new_data[0].length - 1);
                    target.values = reply.new_data;
                    target.format.autofitColumns();
                    newSheet.activate();
                    await context.sync();
                }
            } else {
                resultDiv.className = "status-error";
                resultDiv.innerText = "⚠️ " + reply.message;
            }
        });
    } catch (e) {
        resultDiv.className = "status-error";
        resultDiv.innerText = "❌ Ошибка Excel: " + e.message;
    }
}
