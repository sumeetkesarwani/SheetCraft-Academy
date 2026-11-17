import { GoogleGenAI, Type } from "@google/genai";

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

export async function handler(event) {
  if (event.httpMethod === "OPTIONS") {
    return {
      statusCode: 204,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "Content-Type",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
      },
      body: "",
    };
  }

  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "Method not allowed" }),
    };
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return {
      statusCode: 500,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        error: "Lab Coach backend key is not configured on this environment.",
      }),
    };
  }

  try {
    const payload = JSON.parse(event.body || "{}");
    const question = typeof payload.question === "string" ? payload.question.trim() : "";
    const skillLevel = typeof payload.skillLevel === "string" ? payload.skillLevel : "foundation";
    const activeCell = typeof payload.activeCell === "string" ? payload.activeCell : "E2";
    const sheetSummary = typeof payload.sheetSummary === "string" ? payload.sheetSummary : "";
    const moduleTitle = typeof payload.moduleTitle === "string" ? payload.moduleTitle : "";

    if (!question) {
      return {
        statusCode: 400,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ error: "Please enter a formula, debugging question, or automation task." }),
      };
    }

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

    return {
      statusCode: 200,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*",
      },
      body: JSON.stringify({ coachCard: parsed }),
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unable to reach Lab Coach service.";
    return {
      statusCode: 500,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: message }),
    };
  }
}
