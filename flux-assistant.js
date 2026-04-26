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
 *   User types → Gemini via Netlify Function → JSON action → Validator → Firebase write
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
 *   2. Set GEMINI_API_KEY env var in Netlify
 *   3. Replace this flux-assistant.js in your site
 *   4. Redeploy
 * ════════════════════════════════════════════════════════════════════
 */

(function () {
    "use strict";

    // ───────────────────────────────────────────────────────────────
    //  CONFIG
    // ───────────────────────────────────────────────────────────────

    const FLUX_VERSION = "2.3.0-full-ops";
    const FLUX_MODEL = "gemini-2.5-flash";
    const FLUX_ENDPOINT = "/.netlify/functions/flux";
    const FLUX_DEBUG = false; // Keep internal scanning/JSON/action pulses hidden from users.
    const UNDO_WINDOW_MS = 2500; // 2.5 seconds

    const DATA_PATHS = {
        workOrders: "ees_wo",
        audit: "ees_audit",
        presence: "ees_presence",
        leaves: "ees_leaves",
        sessions: "ees_sessions",
        attendance: "ees_attendance"
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

    function safeFirebaseKey(value) {
        return String(value ?? "")
            .trim()
            .replace(/[.#$/\[\]]/g, "_")
            .slice(0, 120);
    }

    function getWORef(id, childPath = "") {
        const db = getDb();
        if (!db) throw new Error("Firebase database unavailable.");

        const key = safeFirebaseKey(id);
        if (!key) throw new Error("Invalid WO ID.");

        const suffix = childPath ? `/${childPath}` : "";
        return db.ref(`${DATA_PATHS.workOrders}/${key}${suffix}`);
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

        const snap = await getWORef(id).once("value");
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
            .orderByChild("createdAt")
            .limitToLast(limit)
            .once("value");

        const entries = [];
        snap.forEach(child => entries.push(child.val()));

        return entries
            .filter(Boolean)
            .sort((a, b) => Number(b.createdAt || 0) - Number(a.createdAt || 0));
    }


    function formatFluxTime(ms) {
        const n = Number(ms || 0);
        if (!n) return "Not recorded";
        try {
            return new Date(n).toLocaleString("en-GB", {
                day: "2-digit",
                month: "short",
                year: "numeric",
                hour: "2-digit",
                minute: "2-digit"
            });
        } catch {
            return String(n);
        }
    }

    async function snapshotSessions(limit = 25) {
        const db = getDb();
        if (!db) return [];

        const snap = await db
            .ref(DATA_PATHS.sessions)
            .orderByChild("loginAt")
            .limitToLast(limit)
            .once("value");

        const sessions = [];
        snap.forEach(child => {
            const value = child.val() || {};
            sessions.push({
                id: child.key,
                uid: value.uid || "",
                rc: value.rc || "",
                name: value.name || "",
                email: value.email || "",
                role: value.role || "",
                status: value.status || "unknown",
                loginAt: Number(value.loginAt || 0),
                loginAtText: formatFluxTime(value.loginAt),
                logoutAt: Number(value.logoutAt || 0),
                logoutAtText: formatFluxTime(value.logoutAt),
                lastSeenAt: Number(value.lastSeenAt || 0),
                lastSeenAtText: formatFluxTime(value.lastSeenAt)
            });
        });

        return sessions.sort((a, b) => Number(b.loginAt || 0) - Number(a.loginAt || 0));
    }

    async function snapshotAttendance() {
        const db = getDb();
        if (!db) return [];

        const [attendanceSnap, leavesSnap, presenceSnap] = await Promise.all([
            db.ref(DATA_PATHS.attendance).once("value"),
            db.ref(DATA_PATHS.leaves).once("value"),
            db.ref(DATA_PATHS.presence).once("value")
        ]);

        const attendance = attendanceSnap.val() || {};
        const leaves = leavesSnap.val() || {};
        const presence = presenceSnap.val() || {};
        const now = Date.now();

        return WORKERS.map(worker => {
            const rc = String(worker.rc);
            const ownAttendance = attendance[rc] || {};
            const leave = leaves[rc] || null;
            const lastSeen = Number(presence[rc] || 0);
            const online = lastSeen > 0 && (now - lastSeen) < 300000;
            const status = leave?.type && leave.type !== "None"
                ? "On Leave"
                : (ownAttendance.status || (online ? "On Duty" : "Off Duty"));

            return {
                rc,
                name: worker.name,
                status,
                online,
                lastSeenAt: lastSeen,
                lastSeenAtText: formatFluxTime(lastSeen),
                leaveType: leave?.type || "",
                updatedAt: Number(ownAttendance.updatedAt || 0),
                updatedAtText: formatFluxTime(ownAttendance.updatedAt),
                updatedBy: ownAttendance.updatedBy?.name || ownAttendance.updatedBy || ""
            };
        });
    }

    // ───────────────────────────────────────────────────────────────
    //  AUDIT LOG (uses app.js auditLog if present, else writes direct)
    // ───────────────────────────────────────────────────────────────

    async function logFluxAction(action, payload = {}) {
        const combined = { ...payload, source: "flux", fluxVersion: FLUX_VERSION };
        if (typeof window.auditLog === "function") {
            try { await window.auditLog(`FLUX_${action}`, combined); return; } catch {}
        }
        // Fallback: direct write with Firebase server timestamp.
        const db = getDb();
        const auth = getAuth();
        if (!db) return;
        try {
            const auditRef = db.ref(DATA_PATHS.audit).push();

            await auditRef.set({
                action: `FLUX_${action}`,
                payload: combined,
                user: {
                    uid: auth?.currentUser?.uid || "",
                    email: auth?.currentUser?.email || ""
                },
                source: "flux",
                fluxVersion: FLUX_VERSION,
                createdAt: firebase.database.ServerValue.TIMESTAMP
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
            params: { priority: "string?", highPriority: "boolean?", progressBelow: "number?", assigneeRc: "string?" },
            isWrite: false,
            execute: async ({ priority, highPriority, progressBelow, assigneeRc }) => {
                const all = await snapshotWOs(500);
                return all.filter(wo => {
                    if (highPriority && !["Urgent", "Critical"].includes(wo.priority)) return false;
                    if (priority && wo.priority !== priority) return false;
                    if (progressBelow !== undefined && wo.progress >= progressBelow) return false;
                    if (assigneeRc && !wo.assignees.includes(assigneeRc)) return false;
                    return true;
                });
            }
        },

        QUERY_SESSIONS: {
            description: "Read login/session history.",
            params: { rc: "string?", name: "string?", limit: "number?" },
            isWrite: false,
            execute: async ({ rc, name, limit }) => {
                const sessions = await snapshotSessions(Math.min(Number(limit || 25), 100));
                const needle = String(name || "").trim().toLowerCase();
                return sessions.filter(session => {
                    if (rc && String(session.rc) !== String(rc)) return false;
                    if (needle && !String(session.name || "").toLowerCase().includes(needle)) return false;
                    return true;
                });
            }
        },

        QUERY_ATTENDANCE: {
            description: "Read personnel attendance/on-duty/leave status.",
            params: { rc: "string?", status: "string?" },
            isWrite: false,
            execute: async ({ rc, status }) => {
                const attendance = await snapshotAttendance();
                return attendance.filter(row => {
                    if (rc && String(row.rc) !== String(rc)) return false;
                    if (status && String(row.status).toLowerCase() !== String(status).toLowerCase()) return false;
                    return true;
                });
            }
        },

        UPDATE_ATTENDANCE: {
            description: "Update a worker attendance status.",
            params: { rc: "string", status: "On Duty|Off Duty|On Leave|Absent" },
            isWrite: true,
            validate: ({ rc, status }) =>
                workerByRc.has(String(rc)) && ["On Duty", "Off Duty", "On Leave", "Absent"].includes(status),
            execute: async ({ rc, status }) => {
                const db = getDb();
                const key = safeFirebaseKey(rc);
                const ref = db.ref(`${DATA_PATHS.attendance}/${key}`);
                const prevSnap = await ref.once("value");
                const previous = prevSnap.exists() ? prevSnap.val() : null;
                await ref.update({
                    rc: String(rc),
                    name: workerByRc.get(String(rc)) || String(rc),
                    status,
                    updatedAt: firebase.database.ServerValue.TIMESTAMP,
                    updatedBy: {
                        uid: getAuth()?.currentUser?.uid || "",
                        rc: window.currentUser?.rc || "",
                        name: window.currentUser?.name || "",
                        email: getAuth()?.currentUser?.email || ""
                    },
                    source: "flux"
                });
                return { updated: true, previous };
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
                await getWORef(id, "priority").set(priority);
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
                await getWORef(id).update({
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
                await getWORef(id).update({
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
                await getWORef(id, "assignees").set(assignees);
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
                await getWORef(id, "assignees").set(filtered);
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
                await getWORef(id, "tasks").set(tasks);
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
                id: json.id || json.rc,
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
            <div class="flux-undo-msg">✨ Flux updated <b>${escapeHtml(json.id || json.rc || "record")}</b></div>
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
            const baseRef = getWORef(snap.id);

            if (snap.action === "UPDATE_WO_PRIORITY") {
                await getWORef(snap.id, "priority").set(snap.previous.priority);
            } else if (snap.action === "UPDATE_TASK_STATUS" || snap.action === "UPDATE_TASK_PROGRESS" || snap.action === "ADD_TASK_REMARK") {
                const wo = await readWO(snap.id);
                const tasks = [...(wo.tasks || [])];
                if (tasks[snap.taskIndex]) {
                    tasks[snap.taskIndex] = { ...tasks[snap.taskIndex], ...snap.previous };
                    // Recompute overall progress
                    const total = tasks.reduce((a, t) => a + (t.progress || 0), 0);
                    const overall = tasks.length > 0 ? Math.round(total / tasks.length) : 0;
                    await baseRef.update({ tasks, overallProgress: overall });
                }
            } else if (snap.action === "ADD_ASSIGNEE" || snap.action === "REMOVE_ASSIGNEE") {
                await getWORef(snap.id, "assignees").set(snap.previous.assignees || []);
            } else if (snap.action === "UPDATE_ATTENDANCE") {
                const attendanceRef = db.ref(`${DATA_PATHS.attendance}/${safeFirebaseKey(snap.id)}`);
                if (snap.previous) await attendanceRef.set(snap.previous);
                else await attendanceRef.remove();
            }

            await logFluxAction("ACTION_UNDONE", { action: snap.action, id: snap.id });
            addPulse(`↩️ Undid ${snap.action} on ${snap.id}`);
            addMessage("bot", `Undone. ${snap.id} restored.`);
        } catch (err) {
            addPulse(`❌ Undo failed: ${err.message}`);
        }
    }

    // ───────────────────────────────────────────────────────────────
    //  SYSTEM PROMPT — the contract between Flux UI and Gemini
    // ───────────────────────────────────────────────────────────────

    function getFluxTrainingText() {
        const training = window.FLUX_TRAINING;

        if (!training || !Array.isArray(training.examples) || training.examples.length === 0) {
            return "";
        }

        return training.examples
            .map(example => `- ${example}`)
            .join("\n");
    }

    function buildSystemPrompt() {
        const userName = window.currentUser?.name || "Admin";
        const workersList = WORKERS
            .map(w => `  - RC ${w.rc}: ${w.name}`)
            .join("\n");
        const trainingText = getFluxTrainingText();

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

9. QUERY_SESSIONS — read login/logout/session history
   \`\`\`json
   {"action": "QUERY_SESSIONS", "name": "Ahmed", "limit": 10}
   \`\`\`

10. QUERY_ATTENDANCE — read personnel attendance/on-duty/leave status
   \`\`\`json
   {"action": "QUERY_ATTENDANCE", "status": "On Duty"}
   \`\`\`

11. UPDATE_ATTENDANCE — update personnel attendance status
   \`\`\`json
   {"action": "UPDATE_ATTENDANCE", "rc": "7485", "status": "On Duty"}
   \`\`\`

LOGIN AND ATTENDANCE RULES
- If the user asks about login, logout, last online, session history, or who used the system, use QUERY_SESSIONS.
- If the user asks who is on duty, absent, on leave, available, or off duty, use QUERY_ATTENDANCE.
- If the database context does not contain enough session or attendance data, say exactly what signal is missing.

ASSIGNMENT RULES
- When the user says "assign Adam to W00038-EEW", convert Adam Muneef to RC 7485 and use ADD_ASSIGNEE.
- When assigning employees, always write RC numbers, not names.
- Never assign a worker if the name is ambiguous. Ask a short clarification question.
- Treat "urgent WOs" as highPriority=true unless the user explicitly asks for priority exactly Urgent.

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

${trainingText ? `\nADDITIONAL USER PHRASES TO UNDERSTAND\n${trainingText}\n` : ""}

TONE
Concise. Operational. No fluff. You are a tool, not a chatbot.`;
    }

    // ───────────────────────────────────────────────────────────────
    //  GEMINI API CALLER (via Netlify Function)
    // ───────────────────────────────────────────────────────────────

    async function askFluxModel(userMessage, contextData) {
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
    //  JSON EXTRACTOR — pulls action blocks out of Gemini's response
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
        if (!FLUX_DEBUG) return;
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
            // Build a fresh compact operational context for every request.
            const [wos, audits, sessions, attendance] = await Promise.all([
                snapshotWOs(120),
                snapshotAudit(10),
                snapshotSessions(25),
                snapshotAttendance()
            ]);
            const context = {
                workOrders: wos,
                recentAudit: audits,
                recentSessions: sessions,
                attendance
            };

            const reply = await askFluxModel(userMsg, context);

            const actionJson = extractActionJSON(reply);

            if (actionJson) {
                // Gemini wants to DO something
                const spec = ALLOWED_ACTIONS[actionJson.action];
                if (spec && !spec.isWrite) {
                    // Read action — execute and feed back
                    addPulse(`⚡ Reading: ${actionJson.action}`);
                    const result = await processAgentAction(actionJson);
                    // Feed result back to Gemini for final natural-language reply
                    const followup = await askFluxModel(
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
                        // Feed confirmation back so Gemini can explain
                        const followup = await askFluxModel(
                            `[SIGNAL CONFIRMED]\nAction: ${actionJson.action}\nResult: ${JSON.stringify(outcome)}\n\nConfirm to the user in 1-2 sentences.`,
                            null
                        );
                        const clean = stripJsonBlocks(followup) || "Done.";
                        addMessage("bot", clean);
                    } catch (err) {
                        addPulse(`❌ ${err.message}`);
                        const followup = await askFluxModel(
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
            elements.panel.classList.toggle("show", panelOpen);
            elements.button?.classList.toggle("active", panelOpen);
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
                <div class="flux-foot">Powered by Gemini · Operational access · Actions logged</div>
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

        addMessage("bot", "Flux online. I can read WOs, sessions, attendance, audit signals, and execute approved actions.");
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
        window.addEventListener("ees:userchange", () => {
            const authUser = getAuth()?.currentUser;
            if (!authUser) setStatus("Not signed in");
            else setStatus(isAdmin() ? "Ready" : "Read-only (admin required)");
        });
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
