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
            const systemPrompt = `
                You are a witty charismatic recruiter who loves making people laugh
                Your goal is to charm candidates with humor and confidence while staying professional yet casual
                You enjoy personalizing messages to show you have researched their background
                Avoid flattery or generic tones—keep it real sharp and memorable
                Write the message like you are talking to a friend
                Do not use short words such as who is or let us—use full words like who is and let us
                Do not use symbols like exclamation marks commas dashes or parentheses
                Keep it simple and friendly with no extra spaces between lines
            `;
            const userPrompt = `
                Compose a short humorous and hyper-personalized invitation message to ${name} for an interview at ${senderCompany}
                ${senderName} ${senderTitle} at ${senderCompany} is sending this invitation
                Knowing that ${name} works at ${companyName} studied at ${collegeName} and is ${age} years old craft a message that
                Is casual friendly and direct
                Shows genuine research and understanding of ${name}'s background
                Includes a clear call to action
                Avoids formal salutations or closings
                Does not use extra spaces between lines or periods at the end of sentences
                Stays under 100 words
                Does not use symbols like exclamation marks commas dashes or parentheses
                Does not use short words such as who is or let us—use full words like who is and let us
                Focuses on ${name}'s potential and ${senderName}'s expertise without false promises
                Uses humor and personal touches to make it memorable
            `;

            const fullPrompt = `${systemPrompt}\n${userPrompt}`;

            const generationConfig = {
                temperature: 1.0,
                topK: 10,
                maxOutputTokens: 150,
            };

            const result = await model.generateContent(fullPrompt, { generationConfig });
            let communication = result.response.text();

            communication = communication
                .replace(/who's/g, "who is")
                .replace(/let's/g, "let us")
                .replace(/that'll/g, "that will")
                .replace(/wouldn't/g, "would not")
                .replace(/[,!—\-\(\)]/g, "") 
                .replace(/\n\n+/g, "\n") 
                .replace(/\s+/g, " ");
            console.log(communication);
        }
    } catch (error) {
        console.error("Error processing Excel file:", error);
    }
}

generatePersonalizedCommunication("./AI_dataset.xlsx", "Elon Musk", "Hiring Manager", "message", "Google");