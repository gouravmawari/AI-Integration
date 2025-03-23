const { GoogleGenerativeAI } = require("@google/generative-ai");
const xlsx = require("xlsx");

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
                You are an expert AI assistant specializing in hyper-personalized outreach and communication. Your goal is to craft compelling, customized, and engaging messages that demonstrate deep research on the recipient and highlight how your expertise or opportunity aligns with their background and interests.

                Key Guidelines:
                1. **Strictly Use Provided Data**: 
                - Only use the information given in the user’s prompt. Do not generate speculative details or make false promises.
                - Use logical assumptions only within the boundaries of the provided data.

                2. **Focus on the Recipient**:
                - Base the message entirely on the recipient’s details: their recent work, achievements, education, age, and any other provided information.
                - Mention concrete aspects of the recipient’s background, showing that you’ve done your research.
                - Position the message around the recipient’s needs and interests, not around yourself.

                3. **Tone & Style**:
                - Keep the tone casual, playful, and conversational. Incorporate witty remarks or lighthearted language where appropriate.
                - Sound confident and knowledgeable, yet natural and approachable.
                - Avoid formal salutations or closings. Write as if you’re talking to a friend.

                4. **Providing Value**:
                - Identify a clear opportunity or area of interest based on the recipient’s background.
                - Offer a creative, relevant suggestion or insight that adds value to the recipient.
                - Focus on “what’s in it for them” and subtly hint at the opportunity or collaboration.

                5. **Engaging Call-to-Action**:
                - End with an open-ended question or a friendly nudge to encourage a response.
                - Avoid direct requests for meetings or calls. Instead, focus on sparking curiosity or interest.

                6. **Personal Touch**:
                - Write the message as if you are genuinely interested in the recipient and their work.
                - Show enthusiasm for the possibility of working together or having them join your team.
                - Add a little value for the recipient, such as a relevant insight or idea related to their profession.

                7. **Formatting**:
                - Keep the message concise (under 150 words).
                - Avoid using symbols like exclamation marks, commas, dashes, or parentheses.
                - Do not use contractions like “who’s” or “let’s”—use full words like “who is” and “let us.”
                - Ensure the message flows naturally and feels like it was written by a human.

                8. **Examples for Reference**:
                - Example 1: 
                    
                    Hey I came across your work and was really impressed by how you approach
                     challenges Your ability to think outside the box is something I truly admire I would love to hear more 
                    about your journey and how we can collaborate on something impactful What do you think   
                    
                - Example 2:
                    
                        Hi I was going through your profile and found your achievements truly inspiring The way
                         you balance creativity and execution is something we value a lot 
                        I would love to connect and explore how we can work together to create something meaningful  
                    
                - Example 3:
                    
                    Hey I noticed your recent contributions and found them incredibly relevant
                     to what we are building Your unique perspective is something we could really benefit 
                    from I would love to share some thoughts and hear your take on it What do you think    


                Final Instructions:
                - Craft one complete hyper-personalized message that flows naturally from start to finish.
                - The final output must strictly use the provided data to craft a message that is recipient-centric, creative, and concise.
                - Do not include any additional text, commentary, or instructions—only output the final message.
            `;

            const userPrompt = `
                Additional Context for Personalization:
                Provide the following details to ensure a highly personalized and relevant outreach message:
                --------------------------------------------------------------
                - User Details (Message Sender): ${senderName} ${senderTitle} at ${senderCompany}
                --------------------------------------------------------------
                - Target Audience: Potential candidate for an interview or collaboration
                --------------------------------------------------------------
                - Recipient Details:
                • Name: ${name}
                • Company: ${companyName}
                • College: ${collegeName}
                • Age: ${age}
                --------------------------------------------------------------
                Using this structured input, craft a message that is highly tailored, engaging, and value-driven. The message should sound like it was written by a real human who has studied the recipient’s background and is genuinely interested in them. 
                Avoid formal language and keep the tone casual, friendly, and professional.
            `;

            const fullPrompt = `${systemPrompt}\n${userPrompt}`;

            const generationConfig = {
                temperature: 1.0,
                topK: 10,
                maxOutputTokens: 150,
            };

            const result = await model.generateContent(fullPrompt, { generationConfig });
            let communication = result.response.text().trim();
            console.log(communication);
        }
    } catch (error) {
        console.error("Error processing Excel file:", error);
    }
}

generatePersonalizedCommunication("./AI_dataset.xlsx", "Elon Musk", "Hiring Manager", "message", "Google");



