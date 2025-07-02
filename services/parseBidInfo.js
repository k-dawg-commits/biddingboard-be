const { GoogleGenerativeAI } = require("@google/generative-ai")
require("dotenv").config();

const GEMINI_API_KEY = process.env.GEMINI_API_KEY;


const genAi = new GoogleGenerativeAI(GEMINI_API_KEY);
const model = genAi.getGenerativeModel({ model: "gemini-2.0-flash-lite" });

async function parseBidInfo(text) {
    text += process.env.PROMPT
    const result = await model.generateContent(text)
    return result.response.text()
}

module.exports = parseBidInfo;