/*
 * ════════════════════════════════════════════════════════════════════
 *  FLUX COMMAND CENTER
 *  AI-powered write-capable assistant for EES WO Control Database
 *  v2.1.0-gemini — Uses Google Gemini Flash (free tier)
 *
 *  Architecture is provider-agnostic: the Netlify Function adapts this
 *  frontend to whichever LLM you point GEMINI_API_KEY / ANTHROPIC_API_KEY at.
 * ════════════════════════════════════════════════════════════════════
 *
 *  ARCHITECTURE
 *  ────────────
 *   User types → Claude API → JSON action → Validator → Firebase write
 *                                                   ↓
 *                                           Audit log + 2-sec undo
 *
 *  SAFETY RAILS (non-negotiable)
 *  ─────────────────────────────
 *   1. Strict action whitelist (8 actions, no delete/clear/bulk)
 *   2. Every action validated BEFORE execution
 *   3. Every action logged to ees_audit
 *   4. Every write shows 2-second undo banner
 *   5. Admin-only (viewer is blocked client AND server side)
 *
 *  DEPLOY
 *  ──────
 *   1. Deploy netlify/functions/flux.js first
 *   2. Set ANTHROPIC_API_KEY env var in Netlify
 *   3. Replace this flux-assistant.js in your site
 *   4. Redeploy
 * ════════════════════════════════════════════════════════════════════
 */

(function () {
    "use strict";

    // ───────────────────────────────────────────────────────────────
    //  CONFIG
    // ───────────────────────────────────────────────────────────────

    const FLUX_VERSION = "2.1.0-gemini";
    const FLUX_MODEL = "gemini-flash-latest";
    const FLUX_ENDPOINT = "/.netlify/functions/flux";
    const UNDO_WINDOW_MS = 2500; // 2.5 seconds

    const DATA_PATHS = {
        workOrders: "ees_wo",
        audit: "ees_audit",
        presence: "ees_presence",
        leaves: "ees_leaves"
    };

    // Workers list — trimmed down (single source of truth is app.js WORKERS)
    const WORKERS = [
        { rc: "844", name: "Mohamed Abdullah" },
        { rc: "1015", name: "Yoosuf Niyaz" },
        { rc: "4123", name: "Hussain Sunil" },
        { rc: "5992", name: "MD Soharab Hossain" },
        { rc: "6025", name: "Ishag Moosa" },
        { rc: "6031", name: "Zahangir Alam" },
        { rc: "7079", name: "Ali Mafaz" },
        { rc: "7485", name: "Adam Muneef" },
        { rc: "10856", name: "Hassan Uzain Sodig" },
        { rc: "12730", name: "Mohamed Jaisham Ibrahim" },
        { rc: "12866", name: "Selvaraj Rajesh" },
        { rc: "12936", name: "Navedul Hasan" },
        { rc: "12944", name: "MD Reyaz Ansari" },
        { rc: "12965", name: "Shiva Kumar Anil Kumar" },
        { rc: "12966", name: "Shibu Moni" },
        { rc: "13163", name: "Ragunath Ramadoss" }
    ];

    const workerByRc = new Map(WORKERS.map(w => [w.rc, w.name]));

    // ───────────────────────────────────────────────────────────────
    //  STATE
    // ───────────────────────────────────────────────────────────────

    let panelOpen = false;
    let elements = {};
    let conversationHistory = []; // for multi-turn context
    let isProcessing = false;
    let lastActionSnapshot = null; // for undo

    // ───────────────────────────────────────────────────────────────
    //  DOM HELPERS
    // ───────────────────────────────────────────────────────────────

    function escapeHtml(value) {
        return String(value ?? "")
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#39;");
    }

    function nowTime() {
        const d = new Date();
        return d.toTimeString().slice(0, 5);
    }

    // ───────────────────────────────────────────────────────────────
    //  FIREBASE ACCESSORS (Flux never bypasses app rules)
    // ───────────────────────────────────────────────────────────────

    function getDb() {
        if (typeof firebase === "undefined" || !firebase.database) return null;
        try { return firebase.database(); } catch { return null; }
    }

    function getAuth() {
        if (typeof firebase === "undefined" || !firebase.auth) return null;
        try { return firebase.auth(); } catch { return null; }
    }

    function isAdmin() {
        // Delegate to main app's currentUser when available
        if (window.currentUser && window.currentUser.role === "admin") return true;
        return false;
    }

    // Read a work order by ID (returns null if not found)
    async function readWO(id) {
        const db = getDb();
        if (!db) return null;
        const snap = await db.ref(`${DATA_PATHS.workOrders}/${id}`).once("value");
        return snap.val();
    }

    // List of WO summaries for LLM context
    async function snapshotWOs(limit = 60) {
        const db = getDb();
        if (!db) return [];
        const snap = await db.ref(DATA_PATHS.workOrders).once("value");
        const all = Object.values(snap.val() || {});
        // Return compact summaries — too much detail confuses the LLM
        return all.slice(0, limit).map(wo => ({
            id: wo.id,
            asset: wo.asset || "",
            priority: wo.priority || "Minor",
            progress: wo.overallProgress || 0,
            assignees: wo.assignees || [],
            taskCount: Array.isArray(wo.tasks) ? wo.tasks.length : 0
        }));
    }

    // Last N audit log entries
    async function snapshotAudit(limit = 5) {
        const db = getDb();
        if (!db) return [];
        const snap = await db
            .ref(DATA_PATHS.audit)
            .orderByChild("timestamp")
            .limitToLast(limit)
            .once("value");
        const entries = [];
        snap.forEach(child => entries.push(child.val()));
        return entries.reverse();
    }

    // ───────────────────────────────────────────────────────────────
    //  AUDIT LOG (uses app.js auditLog if present, else writes direct)
    // ───────────────────────────────────────────────────────────────

    async function logFluxAction(action, payload = {}) {
        const combined = { ...payload, source: "flux", fluxVersion: FLUX_VERSION };
        if (typeof window.auditLog === "function") {
            try { await window.auditLog(`FLUX_${action}`, combined); return; } catch {}
        }
        // Fallback: direct write
        const db = getDb();
        const auth = getAuth();
        if (!db) return;
        try {
            const key = db.ref(DATA_PATHS.audit).push().key;
            await db.ref(`${DATA_PATHS.audit}/${key}`).set({
                action: `FLUX_${action}`,
                payload: combined,
                timestamp: Date.now(),
                userEmail: auth?.currentUser?.email || "unknown"
            });
        } catch (err) {
            console.warn("Flux audit log failed:", err);
        }
    }

    // ───────────────────────────────────────────────────────────────
    //  ACTION WHITELIST — the ONLY things Flux is allowed to do
    // ───────────────────────────────────────────────────────────────

    const ALLOWED_ACTIONS = {
        QUERY_WO: {
            description: "Read a specific WO's details.",
            params: { id: "string" },
            isWrite: false,
            execute: async ({ id }) => {
                const wo = await readWO(id);
                return wo ? { found: true, wo } : { found: false };
            }
        },

        SEARCH_WO: {
            description: "Find WOs matching a filter.",
            params: { priority: "string?", progressBelow: "number?", assigneeRc: "string?" },
            isWrite: false,
            execute: async ({ priority, progressBelow, assigneeRc }) => {
                const all = await snapshotWOs(500);
                return all.filter(wo => {
                    if (priority && wo.priority !== priority) return false;
                    if (progressBelow !== undefined && wo.progress >= progressBelow) return false;
                    if (assigneeRc && !wo.assignees.includes(assigneeRc)) return false;
                    return true;
                });
            }
        },

        UPDATE_WO_PRIORITY: {
            description: "Change a WO's priority.",
            params: { id: "string", priority: "Minor|Major|Urgent|Critical" },
            isWrite: true,
            validate: ({ priority }) =>
                ["Minor", "Major", "Urgent", "Critical"].includes(priority),
            execute: async ({ id, priority }) => {
                const wo = await readWO(id);
                if (!wo) throw new Error(`WO ${id} not found`);
                const prev = { priority: wo.priority };
                await getDb().ref(`${DATA_PATHS.workOrders}/${id}/priority`).set(priority);
                return { updated: true, previous: prev };
            }
        },

        UPDATE_TASK_STATUS: {
            description: "Set a task's status within a WO.",
            params: {
                id: "string",
                taskIndex: "number",
                status: "Pending|Ongoing|Onhold|Completed|Cancelled"
            },
            isWrite: true,
            validate: ({ status }) =>
                ["Pending", "Ongoing", "Onhold", "Completed", "Cancelled"].includes(status),
            execute: async ({ id, taskIndex, status }) => {
                const wo = await readWO(id);
                if (!wo) throw new Error(`WO ${id} not found`);
                const tasks = Array.isArray(wo.tasks) ? [...wo.tasks] : [];
                if (!tasks[taskIndex]) throw new Error(`Task index ${taskIndex} not found`);
                const prev = { status: tasks[taskIndex].status };
                tasks[taskIndex] = { ...tasks[taskIndex], status };
                // If marking complete and progress < 100, bump to 100
                if (status === "Completed" && (tasks[taskIndex].progress || 0) < 100) {
                    prev.progress = tasks[taskIndex].progress || 0;
                    tasks[taskIndex].progress = 100;
                }
                // Recompute overall
                const total = tasks.reduce((acc, t) => acc + (t.progress || 0), 0);
                const overall = tasks.length > 0 ? Math.round(total / tasks.length) : 0;
                await getDb().ref(`${DATA_PATHS.workOrders}/${id}`).update({
                    tasks,
                    overallProgress: overall
                });
                return { updated: true, previous: prev, newOverall: overall };
            }
        },

        UPDATE_TASK_PROGRESS: {
            description: "Set a task's progress percentage (0-100).",
            params: { id: "string", taskIndex: "number", progress: "number" },
            isWrite: true,
            validate: ({ progress }) =>
                typeof progress === "number" && progress >= 0 && progress <= 100,
            execute: async ({ id, taskIndex, progress }) => {
                const wo = await readWO(id);
                if (!wo) throw new Error(`WO ${id} not found`);
                const tasks = Array.isArray(wo.tasks) ? [...wo.tasks] : [];
                if (!tasks[taskIndex]) throw new Error(`Task index ${taskIndex} not found`);
                const prev = { progress: tasks[taskIndex].progress || 0 };
                tasks[taskIndex] = { ...tasks[taskIndex], progress };
                const total = tasks.reduce((acc, t) => acc + (t.progress || 0), 0);
                const overall = tasks.length > 0 ? Math.round(total / tasks.length) : 0;
                await getDb().ref(`${DATA_PATHS.workOrders}/${id}`).update({
                    tasks,
                    overallProgress: overall
                });
                return { updated: true, previous: prev, newOverall: overall };
            }
        },

        ADD_ASSIGNEE: {
            description: "Assign a worker (by RC number) to a WO.",
            params: { id: "string", rc: "string" },
            isWrite: true,
            validate: ({ rc }) => workerByRc.has(String(rc)),
            execute: async ({ id, rc }) => {
                const wo = await readWO(id);
                if (!wo) throw new Error(`WO ${id} not found`);
                const assignees = Array.isArray(wo.assignees) ? [...wo.assignees] : [];
                if (assignees.includes(rc)) return { updated: false, reason: "already assigned" };
                const prev = { assignees: [...assignees] };
                assignees.push(rc);
                await getDb().ref(`${DATA_PATHS.workOrders}/${id}/assignees`).set(assignees);
                return { updated: true, previous: prev };
            }
        },

        REMOVE_ASSIGNEE: {
            description: "Remove a worker from a WO.",
            params: { id: "string", rc: "string" },
            isWrite: true,
            validate: ({ rc }) => workerByRc.has(String(rc)),
            execute: async ({ id, rc }) => {
                const wo = await readWO(id);
                if (!wo) throw new Error(`WO ${id} not found`);
                const assignees = Array.isArray(wo.assignees) ? [...wo.assignees] : [];
                const prev = { assignees: [...assignees] };
                const filtered = assignees.filter(x => x !== rc);
                if (filtered.length === assignees.length) return { updated: false, reason: "not assigned" };
                await getDb().ref(`${DATA_PATHS.workOrders}/${id}/assignees`).set(filtered);
                return { updated: true, previous: prev };
            }
        },

        ADD_TASK_REMARK: {
            description: "Append a remark to a task (no overwrite).",
            params: { id: "string", taskIndex: "number", remark: "string" },
            isWrite: true,
            validate: ({ remark }) =>
                typeof remark === "string" && remark.length > 0 && remark.length <= 500,
            execute: async ({ id, taskIndex, remark }) => {
                const wo = await readWO(id);
                if (!wo) throw new Error(`WO ${id} not found`);
                const tasks = Array.isArray(wo.tasks) ? [...wo.tasks] : [];
                if (!tasks[taskIndex]) throw new Error(`Task index ${taskIndex} not found`);
                const prev = { remarks: tasks[taskIndex].remarks || "" };
                const existing = tasks[taskIndex].remarks || "";
                const stamp = new Date().toLocaleDateString("en-GB", { day:"2-digit", month:"short" });
                const newRemark = existing
                    ? `${existing}\n[${stamp} Flux] ${remark}`
                    : `[${stamp} Flux] ${remark}`;
                tasks[taskIndex] = { ...tasks[taskIndex], remarks: newRemark };
                await getDb().ref(`${DATA_PATHS.workOrders}/${id}/tasks`).set(tasks);
                return { updated: true, previous: prev };
            }
        }
    };

    // ───────────────────────────────────────────────────────────────
    //  VALIDATOR — runs BEFORE Firebase is touched
    // ───────────────────────────────────────────────────────────────

    function validateAction(json) {
        if (!json || typeof json !== "object") {
            return { ok: false, reason: "Not a valid action object" };
        }
        if (!json.action || typeof json.action !== "string") {
            return { ok: false, reason: "Missing 'action' field" };
        }
        const spec = ALLOWED_ACTIONS[json.action];
        if (!spec) {
            return { ok: false, reason: `Unknown action: ${json.action}` };
        }
        if (spec.isWrite && !isAdmin()) {
            return { ok: false, reason: "Write actions require admin role" };
        }
        if (spec.validate && !spec.validate(json)) {
            return { ok: false, reason: `Invalid parameters for ${json.action}` };
        }
        return { ok: true, spec };
    }

    // ───────────────────────────────────────────────────────────────
    //  EXECUTOR — the only place that touches Firebase for writes
    // ───────────────────────────────────────────────────────────────

    async function processAgentAction(json) {
        const result = validateAction(json);
        if (!result.ok) {
            await logFluxAction("ACTION_REJECTED", { action: json.action, reason: result.reason, input: json });
            throw new Error(`Rejected: ${result.reason}`);
        }
        const { spec } = result;

        addPulse(`⚡ Executing ${json.action}...`);

        let outcome;
        try {
            outcome = await spec.execute(json);
        } catch (err) {
            await logFluxAction("ACTION_FAILED", { action: json.action, error: err.message, input: json });
            throw err;
        }

        // Store undo snapshot for writes
        if (spec.isWrite && outcome && outcome.previous) {
            lastActionSnapshot = {
                action: json.action,
                id: json.id,
                taskIndex: json.taskIndex,
                previous: outcome.previous,
                timestamp: Date.now()
            };
            showUndoBanner(json, outcome);
        }

        await logFluxAction("ACTION_EXECUTED", {
            action: json.action,
            input: json,
            outcome
        });

        return outcome;
    }

    // ───────────────────────────────────────────────────────────────
    //  UNDO SYSTEM — 2.5-second window after every write
    // ───────────────────────────────────────────────────────────────

    function showUndoBanner(json, outcome) {
        // Remove any existing banner
        const old = document.getElementById("flux-undo-banner");
        if (old) old.remove();

        const banner = document.createElement("div");
        banner.id = "flux-undo-banner";
        banner.innerHTML = `
            <div class="flux-undo-msg">✨ Flux updated <b>${escapeHtml(json.id || "record")}</b></div>
            <button class="flux-undo-btn" type="button">Undo</button>
            <div class="flux-undo-bar"><div class="flux-undo-bar-inner"></div></div>
        `;
        document.body.appendChild(banner);

        const undoBtn = banner.querySelector(".flux-undo-btn");
        const bar = banner.querySelector(".flux-undo-bar-inner");

        // Animate bar depleting
        requestAnimationFrame(() => { bar.style.width = "0%"; });

        const timer = setTimeout(() => banner.remove(), UNDO_WINDOW_MS);

        undoBtn.addEventListener("click", async () => {
            clearTimeout(timer);
            banner.remove();
            await performUndo();
        });
    }

    async function performUndo() {
        if (!lastActionSnapshot) {
            addPulse("⚠️ Nothing to undo.");
            return;
        }
        const snap = lastActionSnapshot;
        lastActionSnapshot = null;

        try {
            const db = getDb();
            const path = `${DATA_PATHS.workOrders}/${snap.id}`;

            if (snap.action === "UPDATE_WO_PRIORITY") {
                await db.ref(`${path}/priority`).set(snap.previous.priority);
            } else if (snap.action === "UPDATE_TASK_STATUS" || snap.action === "UPDATE_TASK_PROGRESS" || snap.action === "ADD_TASK_REMARK") {
                const wo = await readWO(snap.id);
                const tasks = [...(wo.tasks || [])];
                if (tasks[snap.taskIndex]) {
                    tasks[snap.taskIndex] = { ...tasks[snap.taskIndex], ...snap.previous };
                    // Recompute overall progress
                    const total = tasks.reduce((a, t) => a + (t.progress || 0), 0);
                    const overall = tasks.length > 0 ? Math.round(total / tasks.length) : 0;
                    await db.ref(path).update({ tasks, overallProgress: overall });
                }
            } else if (snap.action === "ADD_ASSIGNEE" || snap.action === "REMOVE_ASSIGNEE") {
                await db.ref(`${path}/assignees`).set(snap.previous.assignees || []);
            }

            await logFluxAction("ACTION_UNDONE", { action: snap.action, id: snap.id });
            addPulse(`↩️ Undid ${snap.action} on ${snap.id}`);
            addMessage("bot", `Undone. ${snap.id} restored.`);
        } catch (err) {
            addPulse(`❌ Undo failed: ${err.message}`);
        }
    }

    // ───────────────────────────────────────────────────────────────
    //  SYSTEM PROMPT — the contract between Flux UI and Claude
    // ───────────────────────────────────────────────────────────────

    function buildSystemPrompt() {
        const userName = window.currentUser?.name || "Admin";
        const workersList = WORKERS
            .map(w => `  - RC ${w.rc}: ${w.name}`)
            .join("\n");

        return `You are Flux, the operational AI agent for the EES (Electrical & Electronic Services) Work Order Control Database at MTCC.

YOUR USER
You are assisting: ${userName} (role: admin).
Only admins can trigger write actions.

CRITICAL RULES
1. When you need to READ or WRITE data, respond ONLY with a single JSON object wrapped in a code block. No prose before or after it in the same message.
2. When you're giving a natural-language answer to the user, DO NOT include JSON.
3. Your JSON must match one of the ALLOWED ACTIONS below EXACTLY. Do not invent new actions.
4. Never attempt DELETE, CLEAR, BULK operations, or create new WOs. If asked, politely decline.
5. Always use the exact WO IDs and RC numbers given. Never guess or approximate.
6. If you're unsure what the user wants, ASK A CLARIFYING QUESTION before acting.

ALLOWED ACTIONS (case-sensitive)

1. QUERY_WO — read one WO
   \`\`\`json
   {"action": "QUERY_WO", "id": "WO-2847-EEW"}
   \`\`\`

2. SEARCH_WO — filter WOs
   \`\`\`json
   {"action": "SEARCH_WO", "priority": "Urgent", "progressBelow": 100}
   \`\`\`
   All filters optional. priority may be: Minor, Major, Urgent, Critical.

3. UPDATE_WO_PRIORITY
   \`\`\`json
   {"action": "UPDATE_WO_PRIORITY", "id": "WO-2847-EEW", "priority": "Critical"}
   \`\`\`

4. UPDATE_TASK_STATUS
   \`\`\`json
   {"action": "UPDATE_TASK_STATUS", "id": "WO-2847-EEW", "taskIndex": 2, "status": "Completed"}
   \`\`\`
   taskIndex is 0-based. status: Pending|Ongoing|Onhold|Completed|Cancelled.

5. UPDATE_TASK_PROGRESS
   \`\`\`json
   {"action": "UPDATE_TASK_PROGRESS", "id": "WO-2847-EEW", "taskIndex": 0, "progress": 75}
   \`\`\`

6. ADD_ASSIGNEE — assign worker by RC number
   \`\`\`json
   {"action": "ADD_ASSIGNEE", "id": "WO-2847-EEW", "rc": "7485"}
   \`\`\`

7. REMOVE_ASSIGNEE
   \`\`\`json
   {"action": "REMOVE_ASSIGNEE", "id": "WO-2847-EEW", "rc": "7485"}
   \`\`\`

8. ADD_TASK_REMARK — append a note to a task
   \`\`\`json
   {"action": "ADD_TASK_REMARK", "id": "WO-2847-EEW", "taskIndex": 1, "remark": "Parts arrived, resuming work."}
   \`\`\`

WORKER ROSTER
${workersList}

WORKFLOW PATTERN
1. User gives a command in natural language.
2. If ambiguous (e.g. "which task?"), ASK a short clarifying question — do not guess.
3. When you're sure, emit ONE JSON action in a code block.
4. System will execute it and respond with a "Signal Confirmed" or error message.
5. On confirmation, give a SHORT natural-language reply (1-2 sentences) acknowledging the change. No JSON this time.
6. If you need to do multiple things, do them ONE AT A TIME, waiting for each confirmation.

EXAMPLE

User: "Mark task 3 of WO-2847-EEW as complete"
You: \`\`\`json
{"action": "UPDATE_TASK_STATUS", "id": "WO-2847-EEW", "taskIndex": 2, "status": "Completed"}
\`\`\`
System: Signal Confirmed: Task 3 on WO-2847-EEW is now Completed (progress: 100%).
You: Done ✅ Task 3 on WO-2847-EEW is now Completed.

TONE
Concise. Operational. No fluff. You are a tool, not a chatbot.`;
    }

    // ───────────────────────────────────────────────────────────────
    //  CLAUDE API CALLER (via Netlify Function)
    // ───────────────────────────────────────────────────────────────

    async function askClaude(userMessage, contextData) {
        // Build multi-turn context
        conversationHistory.push({
            role: "user",
            content: userMessage
        });

        // Keep only last 10 turns to stay under token limits
        if (conversationHistory.length > 20) {
            conversationHistory = conversationHistory.slice(-20);
        }

        const systemPrompt = buildSystemPrompt();

        // First turn? Prepend context as a "here's the current state" message
        const messagesWithContext = [...conversationHistory];
        if (contextData && conversationHistory.length === 1) {
            messagesWithContext[0] = {
                role: "user",
                content: `[CURRENT STATE SNAPSHOT]\n${JSON.stringify(contextData, null, 2)}\n\n[USER REQUEST]\n${userMessage}`
            };
        }

        const response = await fetch(FLUX_ENDPOINT, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                system: systemPrompt,
                messages: messagesWithContext,
                max_tokens: 1024
            })
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.error || `HTTP ${response.status}`);
        }

        const data = await response.json();
        const text = data.text || "";

        conversationHistory.push({ role: "assistant", content: text });

        return text;
    }

    // ───────────────────────────────────────────────────────────────
    //  JSON EXTRACTOR — pulls action blocks out of Claude's response
    // ───────────────────────────────────────────────────────────────

    function extractActionJSON(text) {
        if (!text || typeof text !== "string") return null;

        // 1. Try fenced code block with ```json ... ```
        const fenced = text.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/);
        if (fenced) {
            try { return JSON.parse(fenced[1]); } catch {}
        }

        // 2. Try any fenced block (Gemini sometimes omits "json" hint)
        const anyFence = text.match(/```\s*(\{[\s\S]*?\})\s*```/);
        if (anyFence) {
            try { return JSON.parse(anyFence[1]); } catch {}
        }

        // 3. Try to find a JSON object containing an "action" key anywhere
        //    (Gemini occasionally skips fences entirely)
        const actionMatch = text.match(/\{[^{}]*"action"\s*:\s*"[A-Z_]+"[^{}]*\}/);
        if (actionMatch) {
            try { return JSON.parse(actionMatch[0]); } catch {}
        }

        // 4. Last resort: greedy search for balanced braces containing "action"
        const allBraces = text.match(/\{[\s\S]*?\}/g);
        if (allBraces) {
            for (const candidate of allBraces) {
                if (candidate.includes('"action"')) {
                    try { return JSON.parse(candidate); } catch {}
                }
            }
        }

        return null;
    }

    function stripJsonBlocks(text) {
        return text.replace(/```(?:json)?[\s\S]*?```/g, "").trim();
    }

    // ───────────────────────────────────────────────────────────────
    //  UI — Terminal-style panel + messages
    // ───────────────────────────────────────────────────────────────

    function addMessage(who, text) {
        if (!elements.messages) return;
        const div = document.createElement("div");
        div.className = `flux-msg flux-msg-${who}`;
        const label = who === "bot" ? "FLUX" : (who === "pulse" ? "▸" : "YOU");
        div.innerHTML = `
            <span class="flux-msg-label">${label}</span>
            <span class="flux-msg-body">${escapeHtml(text)}</span>
            <span class="flux-msg-time">${nowTime()}</span>
        `;
        elements.messages.appendChild(div);
        elements.messages.scrollTop = elements.messages.scrollHeight;
    }

    function addPulse(text) {
        if (!elements.messages) return;
        const div = document.createElement("div");
        div.className = "flux-msg flux-msg-pulse";
        div.innerHTML = `<span class="flux-msg-label">▸</span><span class="flux-msg-body">${escapeHtml(text)}</span>`;
        elements.messages.appendChild(div);
        elements.messages.scrollTop = elements.messages.scrollHeight;
    }

    function setStatus(text) {
        if (elements.status) elements.status.textContent = text;
    }

    // ───────────────────────────────────────────────────────────────
    //  MAIN CONVERSATION LOOP
    // ───────────────────────────────────────────────────────────────

    async function handleSend() {
        if (isProcessing) return;
        const input = elements.input;
        const userMsg = (input.value || "").trim();
        if (!userMsg) return;

        if (!isAdmin()) {
            addMessage("bot", "Flux write features are admin-only. Please sign in as admin.");
            return;
        }

        input.value = "";
        addMessage("user", userMsg);

        isProcessing = true;
        setStatus("Thinking...");
        addPulse("⚡ Scanning database...");

        try {
            // Build context snapshot only on first message of a fresh session
            let context = null;
            if (conversationHistory.length === 0) {
                const [wos, audits] = await Promise.all([
                    snapshotWOs(60),
                    snapshotAudit(5)
                ]);
                context = { workOrders: wos, recentAudit: audits };
                addPulse(`⚡ Loaded ${wos.length} WOs + ${audits.length} audit entries`);
            }

            const reply = await askClaude(userMsg, context);

            const actionJson = extractActionJSON(reply);

            if (actionJson) {
                // Claude wants to DO something
                const spec = ALLOWED_ACTIONS[actionJson.action];
                if (spec && !spec.isWrite) {
                    // Read action — execute and feed back
                    addPulse(`⚡ Reading: ${actionJson.action}`);
                    const result = await processAgentAction(actionJson);
                    // Feed result back to Claude for final natural-language reply
                    const followup = await askClaude(
                        `[SIGNAL CONFIRMED]\nResult: ${JSON.stringify(result).slice(0, 1500)}\n\nNow give the user a concise natural-language summary.`,
                        null
                    );
                    const clean = stripJsonBlocks(followup) || "Read complete.";
                    addMessage("bot", clean);
                } else {
                    // Write action
                    addPulse(`⚡ Writing to Firebase...`);
                    try {
                        const outcome = await processAgentAction(actionJson);
                        addPulse(`✓ Signal confirmed`);
                        // Feed confirmation back so Claude can explain
                        const followup = await askClaude(
                            `[SIGNAL CONFIRMED]\nAction: ${actionJson.action}\nResult: ${JSON.stringify(outcome)}\n\nConfirm to the user in 1-2 sentences.`,
                            null
                        );
                        const clean = stripJsonBlocks(followup) || "Done.";
                        addMessage("bot", clean);
                    } catch (err) {
                        addPulse(`❌ ${err.message}`);
                        const followup = await askClaude(
                            `[SIGNAL FAILED]\nAction: ${actionJson.action}\nError: ${err.message}\n\nExplain the failure briefly.`,
                            null
                        );
                        const clean = stripJsonBlocks(followup) || `Error: ${err.message}`;
                        addMessage("bot", clean);
                    }
                }
            } else {
                // Pure conversation / clarifying question
                const clean = stripJsonBlocks(reply) || reply;
                addMessage("bot", clean);
            }
        } catch (err) {
            addPulse(`❌ ${err.message}`);
            addMessage("bot", `Error: ${err.message}`);
        } finally {
            isProcessing = false;
            setStatus(isAdmin() ? "Ready" : "Read-only (sign in as admin)");
        }
    }

    // ───────────────────────────────────────────────────────────────
    //  UI CONSTRUCTION
    // ───────────────────────────────────────────────────────────────

    function togglePanel(force) {
        panelOpen = typeof force === "boolean" ? force : !panelOpen;
        if (elements.panel) {
            elements.panel.classList.toggle("open", panelOpen);
        }
        if (panelOpen && elements.input) {
            setTimeout(() => elements.input.focus(), 100);
        }
    }

    function createUI() {
        if (document.getElementById("flux-assistant-root")) return;

        const root = document.createElement("div");
        root.id = "flux-assistant-root";
        root.innerHTML = `
            <button class="flux-fab" type="button" aria-label="Open Flux Assistant" title="Flux">⚡</button>
            <section class="flux-panel" aria-label="Flux Assistant">
                <div class="flux-head">
                    <div>
                        <div class="flux-title">▸ FLUX</div>
                        <div class="flux-subtitle">${FLUX_VERSION} · Gemini Agent</div>
                    </div>
                    <button class="flux-close" type="button" aria-label="Close">×</button>
                </div>
                <div class="flux-status">Ready</div>
                <div class="flux-messages"></div>
                <div class="flux-quick">
                    <button type="button" data-flux-q="Show me all Urgent WOs">Urgent</button>
                    <button type="button" data-flux-q="What WOs are below 50% progress?">Lagging</button>
                    <button type="button" data-flux-q="Summarize team workload">Team</button>
                </div>
                <div class="flux-input-row">
                    <input class="flux-input" type="text" placeholder="Type command or question..." autocomplete="off">
                    <button class="flux-send" type="button">→</button>
                </div>
                <div class="flux-foot">Powered by Gemini · Actions logged · 2s undo · Admin only</div>
            </section>
        `;
        document.body.appendChild(root);

        elements = {
            root,
            button: root.querySelector(".flux-fab"),
            panel: root.querySelector(".flux-panel"),
            close: root.querySelector(".flux-close"),
            messages: root.querySelector(".flux-messages"),
            input: root.querySelector(".flux-input"),
            sendButton: root.querySelector(".flux-send"),
            status: root.querySelector(".flux-status")
        };

        elements.button.addEventListener("click", () => togglePanel());
        elements.close.addEventListener("click", () => togglePanel(false));
        elements.sendButton.addEventListener("click", handleSend);
        elements.input.addEventListener("keydown", e => {
            if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSend(); }
        });
        root.querySelectorAll("[data-flux-q]").forEach(btn => {
            btn.addEventListener("click", () => {
                elements.input.value = btn.dataset.fluxQ || "";
                handleSend();
            });
        });

        addMessage("bot", "Flux online (Gemini). I can read and update WOs on your command. Every action is logged and undoable.");
    }

    function initWhenReady() {
        createUI();
        const auth = getAuth();
        if (auth && auth.onAuthStateChanged) {
            auth.onAuthStateChanged(user => {
                if (user) setStatus(isAdmin() ? "Ready" : "Read-only (admin required)");
                else setStatus("Not signed in");
            });
        }
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", initWhenReady);
    } else {
        initWhenReady();
    }

    // Expose for debugging / external integration
    window.FluxCommandCenter = {
        version: FLUX_VERSION,
        processAgentAction,
        validateAction,
        ALLOWED_ACTIONS,
        openPanel: () => togglePanel(true)
    };
})();
