const { GoogleGenerativeAI } = require("@google/generative-ai");
const xlsx = require('xlsx');

async function generatePersonalizedCommunication(excelFilePath, senderName, senderTitle, communicationType) {
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

            let prompt = "";
            let maxTokens = 0;

            switch (communicationType.toLowerCase()) {
                case "email":
                    prompt = `
                        Compose a personalized email to ${name} for an interview invitation.
                        Sender Name: ${senderName}
                        Sender Title: ${senderTitle}
                        Target User Details:
                        - Name: ${name}
                        - Company: ${companyName}
                        - College: ${collegeName}
                        - Age: ${age}
                        Email Tone:
                        - Professional and friendly.
                        - Show that you have done some research on ${name}.
                        - Invite them for an interview.
                        - Write in a human like manner, and avoid robotic responses.
                        - Do not mention that you are an AI.
                        - Do not mention the raw data.
                        Email Content:
                    `;
                    maxTokens = 300;
                    break;
                case "message":
                    prompt = `
                        Compose a short, personalized invitation message to ${name} for an interview.
                        Sender Name: ${senderName}
                        Sender Title: ${senderTitle}
                        Target User Details:
                        - Name: ${name}
                        - Company: ${companyName}
                        - College: ${collegeName}
                        - Age: ${age}
                        Message Tone:
                        - Casual, friendly, and direct.
                        - Show you've done some research on ${name}.
                        - Include a clear call to action (CTA).
                        - Avoid formal salutations or closings.
                        - Keep it under 150 words.
                        - Humorous hyper personalized message.
                        - Focus on prospect and my expertise.
                        - Should have a CTA
                        - Shouldn’t make any false promises
                    `;
                    maxTokens = 150;
                    break;
                default:
                    console.error("Invalid communication type. Use 'email' or 'message'.");
                    continue;
            }

            const generationConfig = {
                temperature: 0.7,
                topK: 10,
                maxOutputTokens: maxTokens,
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

generatePersonalizedCommunication("./AI_dataset.xlsx", "Elon musko", "Hiring Manager", "message");












