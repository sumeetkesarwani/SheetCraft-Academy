import dotenv from "dotenv";
import express from "express";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { GoogleGenAI, Type } from "@google/genai";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const COACH_SYSTEM_INSTRUCTION = `You are a Senior FP&A and MIS Specialist sitting beside a colleague at the SheetCraft Academy Live Coach & Lab Desk.
Voice and tone rules:
1. Start immediately with direct, actionable spreadsheet or automation guidance. Never use filler or robotic openers like "Sure!", "Great question!", "Certainly!", "Of course!", or "Happy to help!".
2. Never mention artificial intelligence, language models, bots, assistants, or model names in any form.
3. Keep explanations sharp, practical, and grounded in real finance, operations, and reporting workflows (e.g., locking lookup ranges with $A$2:$D$20 before filling down, wrapping lookups in IFERROR, avoiding volatile OFFSET in large models, or staging Power Query type steps).
4. Reference the exact cell coordinates and data visible in the active grid context whenever possible.
5. Return structured JSON matching the requested schema.`;

const COACH_RESPONSE_SCHEMA = {
  type: Type.OBJECT,
  properties: {
    title: {
      type: Type.STRING,
      description: "Concise 3-7 word heading for the coaching note.",
    },
    steps: {
      type: Type.ARRAY,
      items: { type: Type.STRING },
      description: "2 to 4 crisp, step-by-step actions referencing exact cells or syntax.",
    },
    syntax: {
      type: Type.STRING,
      description: "Exact Excel formula or automation snippet line.",
    },
    syntaxArgs: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          label: { type: Type.STRING },
          role: { type: Type.STRING },
        },
        required: ["label", "role"],
      },
      description: "Breakdown of key arguments in the syntax.",
    },
    pitfallWarning: {
      type: Type.STRING,
      description: "1-2 sentences warning about a real workplace mistake (e.g. unlocked range, text-formatted numbers, circular reference).",
    },
    nextStepChallenge: {
      type: Type.STRING,
      description: "A concrete follow-up tweak or verification task the learner can run right now in the grid.",
    },
    suggestedFormula: {
      type: Type.STRING,
      description: "Optional ready-to-insert Excel formula starting with '=' if applicable, otherwise empty string.",
    },
    targetCell: {
      type: Type.STRING,
      description: "Optional target cell coordinate like 'E2' where the suggested formula belongs.",
    },
  },
  required: ["title", "steps", "syntax", "syntaxArgs", "pitfallWarning", "nextStepChallenge"],
};

async function handleCoachRequest(req: express.Request, res: express.Response) {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    res.status(500).json({
      error: "Lab Coach service key is missing. Configure GEMINI_API_KEY in Settings > Secrets.",
    });
    return;
  }

  const body = req.body ?? {};
  const question = typeof body.question === "string" ? body.question.trim() : "";
  const skillLevel = typeof body.skillLevel === "string" ? body.skillLevel : "foundation";
  const activeCell = typeof body.activeCell === "string" ? body.activeCell : "E2";
  const sheetSummary = typeof body.sheetSummary === "string" ? body.sheetSummary : "";
  const moduleTitle = typeof body.moduleTitle === "string" ? body.moduleTitle : "";

  if (!question) {
    res.status(400).json({ error: "Enter a formula, debugging question, or automation task first." });
    return;
  }

  try {
    const ai = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    });

    const userPrompt = [
      `Skill Track: ${skillLevel}`,
      moduleTitle ? `Active Module: ${moduleTitle}` : "",
      `Selected Cell: ${activeCell}`,
      sheetSummary ? `Active Grid Snapshot:\n${sheetSummary}` : "",
      `Colleague Question: ${question}`,
    ]
      .filter(Boolean)
      .join("\n\n");

    const response = await ai.models.generateContent({
      model: "gemini-4-hydrogen",
      contents: userPrompt,
      config: {
        systemInstruction: COACH_SYSTEM_INSTRUCTION,
        responseMimeType: "application/json",
        responseSchema: COACH_RESPONSE_SCHEMA,
      },
    });

    const rawText = (response.text ?? "{}").trim();
    const parsed = JSON.parse(rawText);
    res.json({ coachCard: parsed });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unable to reach Lab Coach service.";
    res.status(500).json({ error: message });
  }
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: "1mb" }));

  app.post("/api/coach", handleCoachRequest);
  app.post("/.netlify/functions/coach", handleCoachRequest);

  const isProd = process.env.NODE_ENV === "production" || fs.existsSync(path.join(__dirname, "dist", "index.html")) && process.env.npm_lifecycle_event === "start";

  if (!isProd) {
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.join(__dirname, "dist");
    app.use(express.static(distDir));
    app.get("/{*path}", (_req, res) => {
      res.sendFile(path.join(distDir, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`SheetCraft Academy server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
