/**
 * Netlify Function: flux.js
 * Server-side proxy to Google Gemini API.
 *
 * Why this exists:
 *   - Hides the Gemini API key from the browser (prevents quota theft)
 *   - Provides a stable interface so we can swap models later (Gemini ↔ Claude)
 *     without changing the frontend.
 *
 * Deploy:
 *   1. Place this file at:  netlify/functions/flux.js
 *   2. In Netlify dashboard → Site settings → Environment variables:
 *        GEMINI_API_KEY = AIza...
 *   3. Redeploy the site.
 *
 * The client calls:  fetch("/.netlify/functions/flux", { method: "POST", body: {...} })
 */

const GEMINI_MODEL = "gemini-flash-latest";

exports.handler = async function (event) {
    const corsHeaders = {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "Content-Type",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Content-Type": "application/json"
    };

    if (event.httpMethod === "OPTIONS") {
        return { statusCode: 200, headers: corsHeaders, body: "" };
    }

    if (event.httpMethod !== "POST") {
        return {
            statusCode: 405,
            headers: corsHeaders,
            body: JSON.stringify({ error: "Method not allowed. Use POST." })
        };
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
        return {
            statusCode: 500,
            headers: corsHeaders,
            body: JSON.stringify({
                error: "GEMINI_API_KEY not set in Netlify environment variables."
            })
        };
    }

    let body;
    try {
        body = JSON.parse(event.body || "{}");
    } catch {
        return {
            statusCode: 400,
            headers: corsHeaders,
            body: JSON.stringify({ error: "Invalid JSON in request body." })
        };
    }

    const { system, messages, max_tokens } = body;

    if (!Array.isArray(messages) || messages.length === 0) {
        return {
            statusCode: 400,
            headers: corsHeaders,
            body: JSON.stringify({ error: "messages[] is required." })
        };
    }

    // Convert Claude-style messages to Gemini "contents" format
    //   Claude:  [{ role: "user"|"assistant", content: "..." }]
    //   Gemini:  [{ role: "user"|"model", parts: [{ text: "..." }] }]
    const contents = messages.map(m => ({
        role: m.role === "assistant" ? "model" : "user",
        parts: [{ text: String(m.content || "") }]
    }));

    const geminiBody = {
        contents,
        generationConfig: {
            temperature: 0.3,         // Deterministic for JSON actions
            maxOutputTokens: max_tokens || 1024,
            topP: 0.95
        }
    };

    if (system && typeof system === "string" && system.trim()) {
        geminiBody.systemInstruction = {
            role: "system",
            parts: [{ text: system }]
        };
    }

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

        const response = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "X-goog-api-key": apiKey
            },
            body: JSON.stringify(geminiBody)
        });

        const data = await response.json();

        if (!response.ok) {
            return {
                statusCode: response.status,
                headers: corsHeaders,
                body: JSON.stringify({
                    error: data.error?.message || "Gemini API error",
                    details: data
                })
            };
        }

        let text = "";
        try {
            text = data?.candidates?.[0]?.content?.parts
                ?.map(p => p.text)
                .filter(Boolean)
                .join("") || "";
        } catch {
            text = "";
        }

        if (!text) {
            return {
                statusCode: 502,
                headers: corsHeaders,
                body: JSON.stringify({
                    error: "Gemini returned no text",
                    stop_reason: data?.candidates?.[0]?.finishReason || "unknown",
                    details: data
                })
            };
        }

        return {
            statusCode: 200,
            headers: corsHeaders,
            body: JSON.stringify({
                text,
                usage: data.usageMetadata || null,
                stop_reason: data?.candidates?.[0]?.finishReason || null
            })
        };
    } catch (err) {
        return {
            statusCode: 500,
            headers: corsHeaders,
            body: JSON.stringify({
                error: "Proxy failed to reach Gemini API",
                details: err.message
            })
        };
    }
};
