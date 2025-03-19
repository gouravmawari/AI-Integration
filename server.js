const { GoogleGenerativeAI } = require("@google/generative-ai");
const xlsx = require('xlsx');

async function generatePersonalizedCommunication(excelFilePath, senderName, senderTitle, communicationType, senderCompany) {
    const genAI = new GoogleGenerativeAI(""); 
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

    try {
        const workbook = xlsx.readFile(excelFilePath);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = xlsx.utils.sheet_to_json(sheet);

        for (const userData of jsonData) {
            const { name, companyName, collegeName, age } = userData;

            if (!name || !companyName || !collegeName || !age) {
                console.error("Missing data for user:", userData);
                continue;
            }

            const prompt = `
                Compose a short, humorous, and hyper-personalized invitation message to ${name} for an interview at ${senderCompany}. 
                ${senderName}, ${senderTitle} at ${senderCompany}, is extending this invitation. 
                Knowing that ${name} works at ${companyName}, studied at ${collegeName}, and is ${age} years old, craft a message that:
                - Is casual, friendly, and direct.
                - Shows genuine research and understanding of ${name}'s background.
                - Includes a clear call to action (CTA).
                - Avoids formal salutations or closings.
                - Stays under 150 words.
                - Focuses on ${name}'s potential and ${senderName}'s expertise without making false promises.
                - Uses humor and personal touches to make it memorable.
            `;

            const generationConfig = {
                temperature: 1.0,
                topK: 10,
                maxOutputTokens: 150,
            };

            const result = await model.generateContent(prompt, { generationConfig });
            const communication = result.response.text();

            console.log(`\n${communicationType.charAt(0).toUpperCase() + communicationType.slice(1)} to: ${name}\n`);
            console.log(communication);
        }
    } catch (error) {
        console.error("Error processing Excel file:", error);
    }
}

generatePersonalizedCommunication("./AI_dataset.xlsx", "Elon Musk", "Hiring Manager", "message", "Google");











