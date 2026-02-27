import { GoogleGenAI, Type } from "@google/genai";
import { Good } from "../types";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

export async function parseDocumentWithAI(file: File): Promise<Good[]> {
  const base64Data = await fileToBase64(file);
  
  const prompt = `
    Extract a list of goods/items from this document for a logistics loading system.
    For each item, identify:
    - Name (名称)
    - Weight per unit in kg (单件重量)
    - Volume per unit in m3 (单件体积)
    - Quantity (数量)

    If weight or volume is missing, try to estimate based on the item name or use 0.
    Return the data as a JSON array of objects.
  `;

  const response = await ai.models.generateContent({
    model: "gemini-2.0-flash-exp",
    contents: [
      {
        parts: [
          {
            inlineData: {
              mimeType: file.type,
              data: base64Data,
            },
          },
          { text: prompt },
        ],
      },
    ],
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            name: { type: Type.STRING },
            weight: { type: Type.NUMBER },
            volume: { type: Type.NUMBER },
            quantity: { type: Type.NUMBER },
          },
          required: ["name", "weight", "volume", "quantity"],
        },
      },
    },
  });

  try {
    const parsed = JSON.parse(response.text || "[]");
    return parsed.map((item: any) => ({
      id: Math.random().toString(36).substr(2, 9),
      name: item.name || "未知物品",
      weight: item.weight || 0,
      volume: item.volume || 0,
      quantity: item.quantity || 1,
    }));
  } catch (e) {
    console.error("AI Parsing Error:", e);
    return [];
  }
}

function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      const base64String = (reader.result as string).split(",")[1];
      resolve(base64String);
    };
    reader.onerror = (error) => reject(error);
  });
}
