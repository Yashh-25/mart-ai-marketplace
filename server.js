
import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import ExcelJS from "exceljs";
import { GoogleGenAI } from "@google/genai";
import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;
const __dirname = path.dirname(fileURLToPath(import.meta.url));

const EXCEL_FILE = path.join(__dirname, "marketplace.xlsx");

app.use(cors());
app.use(express.json({ limit: "20kb" }));
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

const ai = process.env.GEMINI_API_KEY
  ? new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY })
  : null;

// Read an Excel worksheet and convert rows to objects.
async function readSheet(sheetName) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE);

  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) return [];

  const headers = sheet.getRow(1).values
    .slice(1)
    .map(value => String(value ?? "").trim());

  const records = [];

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const record = {};

    headers.forEach((header, index) => {
      const cell = row.getCell(index + 1).value;

      record[header] =
        cell && typeof cell === "object" && "text" in cell
          ? cell.text
          : cell ?? "";
    });

    if (Object.values(record).some(value => value !== "")) {
      records.push(record);
    }
  });

  return records;
}

// Simple relevance search against Excel records.
function findRelevant(records, question, limit = 8) {
  const words = question.toLowerCase()
    .split(/[^a-z0-9₹.]+/)
    .filter(word => word.length > 1);

  return records
    .map(record => {
      const text = Object.values(record).join(" ").toLowerCase();

      const score = words.reduce(
        (total, word) => total + (text.includes(word) ? 1 : 0),
        0
      );

      return { record, score };
    })
    .filter(item => item.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, limit)
    .map(item => item.record);
}

// Health check.
app.get("/api/health", (req, res) => {
  res.json({
    status: "ok",
    message: "MART backend is running"
  });
});

// Chat endpoint for the frontend.
app.post("/api/chat", async (req, res) => {
  try {
    const message = req.body?.message?.trim();

    if (!message) {
      return res.status(400).json({
        error: "Please enter a question."
      });
    }

    if (message.length > 1500) {
      return res.status(400).json({
        error: "Your question is too long. Please use 1500 characters or fewer."
      });
    }

    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(500).json({
        error: "marketplace.xlsx was not found in the project folder."
      });
    }

    if (!ai) {
      return res.status(500).json({
        error: "Gemini API key is missing. Add GEMINI_API_KEY to your .env file."
      });
    }

    const [products, suppliers, buyers] = await Promise.all([
      readSheet("Products"),
      readSheet("Suppliers"),
      readSheet("Buyers")
    ]);

    const relevantProducts = findRelevant(products, message);
    const relevantSuppliers = findRelevant(suppliers, message);
    const relevantBuyers = findRelevant(buyers, message, 4);

    // Only provide the model with relevant records.
    const context = {
      products: relevantProducts,
      suppliers: relevantSuppliers,
      buyers: relevantBuyers
    };

    const prompt = `
You are MART AI, a helpful assistant for a B2B product
and supplier marketplace.

Answer the buyer's question using the supplied Excel data.

Rules:
- Do not invent products, suppliers, prices, stock, or contact details.
- Prices and supplier information must come from the supplied data.
- Clearly say when matching information is unavailable.
- Compare suitable options when possible.
- Respect the buyer's stated budget, location, and quantity.
- Treat all records as demo data unless independently verified.
- Do not expose buyer personal details or buyer records to other users.
- Keep the answer clear, practical, and concise.

Excel data:
${JSON.stringify(context, null, 2)}

Buyer's question:
${message}
`;

    const result = await ai.models.generateContent({
      model: process.env.GEMINI_MODEL || "gemini-2.5-flash",
      contents: prompt
    });

    res.json({
      reply: result.text || "I couldn't generate an answer. Please try again."
    });

  } catch (error) {
    console.error("Chat error:", error.message);

    res.status(500).json({
      error: "The AI request failed. Check the server terminal and try again."
    });
  }
});

app.listen(PORT, () => {
  console.log(`MART backend running at http://localhost:${PORT}`);
});
