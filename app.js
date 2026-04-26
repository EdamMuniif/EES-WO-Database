// ── STEP 1: FIREBASE CONFIGURATION ──────────────────────────────
const firebaseConfig = {
    apiKey: "AIzaSyCU0sZN9LRTdXjI62J2v-vtML-wYMfF80c",
    authDomain: "ees-wo-database.firebaseapp.com",
    databaseURL: "https://ees-wo-database-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "ees-wo-database",
    storageBucket: "ees-wo-database.firebasestorage.app",
    messagingSenderId: "979788754009",
    appId: "1:979788754009:web:a3334a5dc6d036b23d5548",
    measurementId: "G-8VEQ08N504"
};
// ─────────────────────────────────────────────────────────────────

// ── Boot dependency checks ───────────────────────────────────────
function showFatalStartupError(message) {
    const overlay = document.getElementById('loading-overlay');
    const msg = document.getElementById('loading-msg');
    if (msg) msg.textContent = message;
    if (overlay) overlay.style.display = 'flex';
    const root = document.getElementById('app-root');
    if (root) {
        root.innerHTML = `<div class="login-wrap"><div class="login-card"><div class="login-title">Startup Error</div><div class="login-err" style="text-align:center;margin-top:18px;">${escapeHtml(message)}</div></div></div>`;
    }
}

if (!window.firebase || !firebase.initializeApp || !firebase.database || !firebase.auth) {
    showFatalStartupError("Firebase failed to load. Check your connection and refresh.");
    throw new Error("Firebase SDK unavailable");
}

firebase.initializeApp(firebaseConfig);
const db   = firebase.database();
const auth = firebase.auth();

// ── User role map ─────────────────────────────────────────────────
const USER_ROLES = {
    'adam.muneef@mtcc.com.mv':    { rc:'7485', name:'Adam Muneef',    designation:'Senior Electrician',  role:'admin'  },
    'ahmed.miushaan@mtcc.com.mv': { rc:'5794', name:'Ahmed Miushaan', designation:'Project Coordinator', role:'viewer' }
};

const PRI_CFG = {
    "Minor":    { color: "#8e8e93" },
    "Major":    { color: "#ffcc00" },
    "Urgent":   { color: "#ff9500" },
    "Critical": { color: "#ff3b30" }
};

const WORKERS = [
    { rc:"844",   name:"Mohamed Abdullah",        designation:"Senior Technical Manager", phone:"",        email:"" },
    { rc:"1015",  name:"Yoosuf Niyaz",            designation:"Outpost Service Manager",  phone:"",        email:"" },
    { rc:"4123",  name:"Hussain Sunil",            designation:"Technical Assistant",      phone:"",        email:"" },
    { rc:"5992",  name:"MD Soharab Hossain",       designation:"Electrician",              phone:"",        email:"" },
    { rc:"6025",  name:"Ishag Moosa",              designation:"Senior Electrician",       phone:"",        email:"" },
    { rc:"6031",  name:"Zahangir Alam",            designation:"Senior Electrician",       phone:"",        email:"" },
    { rc:"7079",  name:"Ali Mafaz",                designation:"Technical Assistant",      phone:"",        email:"" },
    { rc:"7485",  name:"Adam Muneef",              designation:"Senior Electrician",       phone:"9515707", email:"adam.muneef@mtcc.com.mv" },
    { rc:"10856", name:"Hassan Uzain Sodig",       designation:"Electrician",              phone:"",        email:"" },
    { rc:"12730", name:"Mohamed Jaisham Ibrahim",  designation:"Technical Assistant",      phone:"",        email:"" },
    { rc:"12866", name:"Selvaraj Rajesh",          designation:"Electrician",              phone:"",        email:"" },
    { rc:"12936", name:"Navedul Hasan",            designation:"Electrician",              phone:"",        email:"" },
    { rc:"12944", name:"MD Reyaz Ansari",          designation:"Electrician",              phone:"",        email:"" },
    { rc:"12965", name:"Shiva Kumar Anil Kumar",   designation:"AC Technician",            phone:"",        email:"" },
    { rc:"12966", name:"Shibu Moni",               designation:"AC Technician",            phone:"",        email:"" },
    { rc:"13163", name:"Ragunath Ramadoss",        designation:"AC Technician",            phone:"",        email:"" }
];

const ini = n => (n||"").split(" ").map(x=>x[0]).join("").slice(0,2).toUpperCase();
const pageLabels = {
    dashboard: "Dashboard",
    orders: "Active Orders",
    urgent: "Urgent & Critical",
    ongoing: "Ongoing WOs",
    completed: "WO Completed",
    upload: "Manage Data",
    workers: "EES Team"
};


// ── Date helpers ──────────────────────────────────────────────────
function formatForInput(val) {
    if(!val || String(val).trim()===""||String(val).trim()==="-") return "";
    let s=String(val).trim(), y,m,d;
    try {
        if(s.includes('/')) { let p=s.split('/'); if(p.length===3){d=+p[0];m=+p[1];y=+p[2];if(m>12&&d<=12){let t=d;d=m;m=t;}if(y<100)y+=2000;} }
        else if(s.includes('-')) { let p=s.split('-'); if(p.length===3){if(p[0].length===4){y=+p[0];m=+p[1];d=+p[2];}else{d=+p[0];y=+p[2];if(y<100)y+=2000;const mn=["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];m=mn.findIndex(x=>p[1].toUpperCase().includes(x))+1;if(m===0)m=+p[1];}}}
        else if(!isNaN(parseFloat(s))&&parseFloat(s)>20000){let o=new Date(Math.round((parseFloat(s)-25569)*86400*1000));y=o.getUTCFullYear();m=o.getUTCMonth()+1;d=o.getUTCDate();}
        else{let o=new Date(s);if(!isNaN(o.getTime())){y=o.getFullYear();m=o.getMonth()+1;d=o.getDate();}}
        if(y&&m&&d&&y>=2000&&y<=2100) return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    } catch(e) {}
    return "";
}

function formatDateNice(ds) {
    if(!ds||ds.includes("1970")||ds.includes("1899")) return "-";
    const p=ds.split('-'); if(p.length!==3) return ds;
    const y=+p[0],mo=+p[1]-1,d=+p[2]; if(y<2000) return "-";
    const mn=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    const sx=["th","st","nd","rd"][(d%10>3||Math.floor(d%100/10)===1)?0:d%10];
    return `${d}${sx} ${mn[mo]} ${y}`;
}

const isBadDate = ds => !ds||ds===""||ds.includes("1970")||ds.includes("1899");

// ── Safety helpers ────────────────────────────────────────────────
function clampNumber(value, min, max) {
    const num = Number(value);
    if (!Number.isFinite(num)) return min;
    return Math.min(max, Math.max(min, num));
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('\"', "&quot;")
        .replaceAll("'", "&#039;");
}

function html(value) {
    return escapeHtml(value);
}

function attr(value) {
    return escapeHtml(value);
}

function jsArg(value) {
    // Safe for inline onclick handlers inside double-quoted HTML attributes.
    return escapeHtml(JSON.stringify(String(value ?? "")));
}

function safeFirebaseKey(value) {
    return String(value ?? "")
        .trim()
        .replace(/[.#$/\[\]]/g, "_")
        .slice(0, 120);
}

// ── Day 3: Role / permission helpers ─────────────────────────────
function normalizeRole(role) {
    return role === "admin" ? "admin" : "viewer";
}

function isAdminUser(profile = currentUser) {
    return profile?.role === "admin";
}

function isViewerUser(profile = currentUser) {
    return profile?.role === "viewer";
}

function requireAdmin(actionName = "perform this action") {
    if (isAdminUser()) return true;
    showToast(`❌ Only admins can ${actionName}.`, true);
    return false;
}

function assertAdmin(actionName = "perform this action") {
    if (!isAdminUser()) {
        throw new Error(`Only admins can ${actionName}.`);
    }
}

async function loadUserProfile(firebaseUser) {
    const email = String(firebaseUser?.email || "").trim().toLowerCase();
    const fallback = USER_ROLES[email] || null;
    const uid = safeFirebaseKey(firebaseUser?.uid || "");

    // Prefer server-side role data if it exists. Fallback keeps the app usable
    // while you create /ees_roles/<uid> records and before strict rules are deployed.
    if (uid) {
        try {
            const snap = await db.ref(`ees_roles/${uid}`).once("value");
            const roleData = snap.val();

            if (roleData && roleData.active !== false) {
                return {
                    uid,
                    email,
                    rc: String(roleData.rc || fallback?.rc || "").trim(),
                    name: String(roleData.name || fallback?.name || email).trim(),
                    designation: String(roleData.designation || fallback?.designation || "User").trim(),
                    role: normalizeRole(roleData.role || fallback?.role)
                };
            }
        } catch (error) {
            console.warn("Could not read server role profile; using local fallback if available.", error);
        }
    }

    return fallback ? { ...fallback, uid, email, role: normalizeRole(fallback.role) } : null;
}

async function auditLog(action, payload = {}) {
    // Audit records use Firebase server time only.
    // No client Date.now() is used for audit ordering.
    if (!currentUser) return null;

    try {
        const auditRef = db.ref("ees_audit").push();

        await auditRef.set({
            action: String(action || "UNKNOWN_ACTION"),
            payload: payload || {},
            user: {
                uid: auth.currentUser?.uid || currentUser.uid || "",
                rc: currentUser.rc || "",
                name: currentUser.name || "",
                email: currentUser.email || auth.currentUser?.email || ""
            },
            userAgent: navigator.userAgent || "",
            createdAt: firebase.database.ServerValue.TIMESTAMP
        });

        return auditRef.key;
    } catch (error) {
        console.warn("Audit log failed", { action, error });
        return null;
    }
}

window.auditLog = auditLog;


// ── Day 23: Login/session + attendance records ─────────────────────
let activeSessionRef = null;
let activeSessionUid = "";

function currentActorPayload() {
    return {
        uid: auth.currentUser?.uid || currentUser?.uid || "",
        rc: currentUser?.rc || "",
        name: currentUser?.name || "",
        email: currentUser?.email || auth.currentUser?.email || ""
    };
}

async function startUserSession() {
    if (!auth.currentUser || !currentUser || !db) return;

    const uid = auth.currentUser.uid;
    const rc = safeFirebaseKey(currentUser.rc || "");
    if (!uid || !rc) return;

    // Prevent duplicate session rows if auth state and manual login both fire.
    if (activeSessionRef && activeSessionUid === uid) return;

    const sessionRef = db.ref("ees_sessions").push();
    activeSessionRef = sessionRef;
    activeSessionUid = uid;

    const actor = currentActorPayload();
    const sessionData = {
        uid,
        rc: currentUser.rc || "",
        name: currentUser.name || "",
        email: currentUser.email || auth.currentUser.email || "",
        role: currentUser.role || "",
        status: "online",
        userAgent: navigator.userAgent || "",
        loginAt: firebase.database.ServerValue.TIMESTAMP,
        lastSeenAt: firebase.database.ServerValue.TIMESTAMP
    };

    await sessionRef.set(sessionData);

    await sessionRef.onDisconnect().update({
        status: "offline",
        logoutAt: firebase.database.ServerValue.TIMESTAMP,
        lastSeenAt: firebase.database.ServerValue.TIMESTAMP
    });

    const presenceRef = db.ref(`ees_presence/${rc}`);
    await presenceRef.set(firebase.database.ServerValue.TIMESTAMP);
    presenceRef.onDisconnect().set(0);

    const attendanceRef = db.ref(`ees_attendance/${rc}`);
    await attendanceRef.update({
        rc: currentUser.rc || "",
        name: currentUser.name || "",
        status: workerLeaves[currentUser.rc]?.type ? "On Leave" : "On Duty",
        source: "login",
        updatedAt: firebase.database.ServerValue.TIMESTAMP,
        updatedBy: actor
    });

    attendanceRef.onDisconnect().update({
        status: "Off Duty",
        source: "disconnect",
        updatedAt: firebase.database.ServerValue.TIMESTAMP,
        updatedBy: {
            uid,
            rc: currentUser.rc || "",
            name: currentUser.name || "",
            email: currentUser.email || auth.currentUser.email || ""
        }
    });

    await auditLog("SESSION_LOGIN", { rc: currentUser.rc || "", name: currentUser.name || "" });
}

async function endUserSession() {
    const actor = currentActorPayload();

    try {
        if (activeSessionRef) {
            await activeSessionRef.update({
                status: "offline",
                logoutAt: firebase.database.ServerValue.TIMESTAMP,
                lastSeenAt: firebase.database.ServerValue.TIMESTAMP
            });
        }

        if (currentUser?.rc) {
            const rc = safeFirebaseKey(currentUser.rc);
            await db.ref(`ees_presence/${rc}`).set(0);
            await db.ref(`ees_attendance/${rc}`).update({
                rc: currentUser.rc || "",
                name: currentUser.name || "",
                status: workerLeaves[currentUser.rc]?.type ? "On Leave" : "Off Duty",
                source: "logout",
                updatedAt: firebase.database.ServerValue.TIMESTAMP,
                updatedBy: actor
            });
            await auditLog("SESSION_LOGOUT", { rc: currentUser.rc || "", name: currentUser.name || "" });
        }
    } catch (error) {
        console.warn("Session logout update failed:", error);
    } finally {
        activeSessionRef = null;
        activeSessionUid = "";
    }
}

async function backupOperationalData(reason = "manual_backup") {
    assertAdmin("create database backups");

    const [woSnap, leavesSnap, metaSnap] = await Promise.all([
        db.ref("ees_wo").once("value"),
        db.ref("ees_leaves").once("value"),
        db.ref("ees_meta").once("value")
    ]);

    const backupKey = `${Date.now()}_${safeFirebaseKey(currentUser?.rc || "admin")}`;
    const backupPath = `ees_backups/${backupKey}`;

    const woData = woSnap.val() || {};
    const leavesData = leavesSnap.val() || {};
    const metaData = metaSnap.val() || {};

    await db.ref(backupPath).set({
        reason,
        createdAt: firebase.database.ServerValue.TIMESTAMP,
        createdBy: {
            uid: auth.currentUser?.uid || currentUser.uid || "",
            rc: currentUser.rc || "",
            name: currentUser.name || "",
            email: currentUser.email || auth.currentUser?.email || ""
        },
        counts: {
            workOrders: Object.keys(woData).length,
            leaves: Object.keys(leavesData).length,
            metaKeys: Object.keys(metaData).length
        },
        data: {
            ees_wo: woData,
            ees_leaves: leavesData,
            ees_meta: metaData
        }
    });

    return backupPath;
}

function normalizePriority(priority) {
    return PRI_CFG[priority] ? priority : "Minor";
}

function normalizeStatus(status) {
    const allowed = new Set(["Pending", "Ongoing", "Onhold", "Completed", "Cancelled"]);
    return allowed.has(status) ? status : "Pending";
}

function normalizeTask(task) {
    return {
        taskNo: String(task?.taskNo ?? "").trim() || "Task",
        details: String(task?.details ?? "").trim(),
        status: normalizeStatus(task?.status),
        startDate: formatForInput(task?.startDate) || "",
        progress: clampNumber(task?.progress, 0, 100),
        completeDate: formatForInput(task?.completeDate) || "",
        remarks: String(task?.remarks ?? "").trim()
    };
}

function normalizeWorkOrder(wo) {
    if (!wo || typeof wo !== "object") throw new Error("Invalid work order object.");
    const id = String(wo.id ?? "").trim();
    if (!id) throw new Error("Work order ID is required.");
    const tasks = Array.isArray(wo.tasks) ? wo.tasks.map(normalizeTask) : [];
    let totalProgress = 0;
    tasks.forEach(task => { totalProgress += task.progress; });
    return {
        id,
        asset: String(wo.asset ?? "TBA").trim() || "TBA",
        date: formatForInput(wo.date) || "",
        sr: String(wo.sr ?? "").trim(),
        svo: String(wo.svo ?? "").trim(),
        priority: normalizePriority(wo.priority),
        overallProgress: tasks.length ? Math.round(totalProgress / tasks.length) : 0,
        assignees: Array.isArray(wo.assignees) ? wo.assignees.map(x => String(x).trim()).filter(Boolean) : [],
        tasks
    };
}

function getFirstWorksheet(workbook) {
    if (!workbook || !Array.isArray(workbook.SheetNames) || workbook.SheetNames.length === 0) {
        throw new Error("No sheets found in this Excel file.");
    }
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    if (!worksheet) {
        throw new Error("The first worksheet could not be read.");
    }
    return worksheet;
}

function isSupportedExcelFile(file) {
    if (!file) return false;
    return /\.(xlsx|xls)$/i.test(file.name || "");
}

function normalizeHeaderKey(key) {
    return String(key || "").trim().replace(/\s+/g, " ").toUpperCase();
}

function findHeader(headers, matcher, fallback = "") {
    return headers.find(header => matcher(normalizeHeaderKey(header))) || fallback;
}

function getValidatedExcelRows(worksheet) {
    let rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    // Some Scope Master sheets have the real header around row 5.
    if (rows.length > 0 && !("WO NUMBER" in rows[0]) && !("WO ID" in rows[0])) {
        rows = XLSX.utils.sheet_to_json(worksheet, { range: 4, defval: "" });
    }

    if (!rows.length) {
        throw new Error("No data rows found in the selected Excel sheet.");
    }

    const headers = Object.keys(rows[0] || {});
    const normalizedHeaders = headers.map(normalizeHeaderKey);
    const hasWorkOrderColumn = normalizedHeaders.includes("WO NUMBER") || normalizedHeaders.includes("WO ID");

    if (!hasWorkOrderColumn) {
        throw new Error("Missing required column: WO NUMBER or WO ID.");
    }

    if (!normalizedHeaders.includes("UNIT")) {
        throw new Error("Missing required column: UNIT.");
    }

    return rows;
}

function getExcelImportColumns(rows) {
    const headers = Object.keys(rows[0] || {});
    return {
        startDate: findHeader(headers, h => h.includes("START DATE"), "START DATE"),
        progress: findHeader(headers, h => h.includes("PROGRESS"), " PROGRESS\n%"),
        status: findHeader(headers, h => h === "STATUS", "STATUS"),
        completedDate: findHeader(headers, h => h === "COMPLETED DATE", "COMPLETED DATE"),
        remarks: findHeader(headers, h => h.includes("REMARKS"), "WORKSHOPS REMARKS"),
        details: findHeader(headers, h => h.includes("CAPTAL") || h === "TASK DETAILS", "TASK DETAILS")
    };
}

function hasUsableTaskData(row, columns) {
    return Boolean(
        String(row[columns.details] || row["TASK DETAILS"] || "").trim() ||
        String(row["TASK NO:"] || "").trim() ||
        String(row[columns.status] || "").trim() ||
        String(row[columns.progress] || "").trim()
    );
}

function parseProgress(value) {
    let progress = 0;
    if (typeof value === "number") progress = value;
    else if (typeof value === "string") progress = parseFloat(value.replace("%", "")) || 0;
    if (progress > 0 && progress <= 1) progress *= 100;
    return Math.round(clampNumber(progress, 0, 100));
}

function parseStatus(value) {
    const rawStatus = String(value || "Pending").trim().toLowerCase();
    if (rawStatus.includes("hold") || rawStatus.includes("material") || rawStatus.includes("delay")) return "Onhold";
    if (rawStatus.includes("ongo") || rawStatus.includes("on-go") || rawStatus.includes("progress")) return "Ongoing";
    if (rawStatus.includes("comp")) return "Completed";
    if (rawStatus.includes("canc")) return "Cancelled";
    return "Pending";
}

// ── Day 4: Excel import indexes + presence helpers ───────────────
function buildOrderIndex(orderList = orders) {
    return new Map(
        orderList
            .map(order => [String(order?.id || "").trim(), order])
            .filter(([id]) => Boolean(id))
    );
}

function getCachedTaskIndex(workOrder, cache) {
    const workOrderId = String(workOrder?.id || "").trim();
    if (!workOrderId) return new Map();

    if (!cache.has(workOrderId)) {
        const taskIndex = new Map(
            (Array.isArray(workOrder.tasks) ? workOrder.tasks : [])
                .map(task => [String(task?.taskNo || "").trim(), task])
                .filter(([taskNo]) => Boolean(taskNo))
        );
        cache.set(workOrderId, taskIndex);
    }

    return cache.get(workOrderId);
}

function makeImportedTask(row, columns, taskNo) {
    return normalizeTask({
        taskNo,
        details: String(row[columns.details] || row["TASK DETAILS"] || "").trim(),
        status: parseStatus(row[columns.status]),
        startDate: formatForInput(row[columns.startDate]),
        progress: parseProgress(row[columns.progress]),
        completeDate: formatForInput(row[columns.completedDate]),
        remarks: String(row[columns.remarks] || "").trim()
    });
}

function mergeImportedTask(existingTask, incomingTask, sourceRow, columns) {
    let changed = false;

    if (incomingTask.startDate && (isBadDate(existingTask.startDate) || existingTask.startDate !== incomingTask.startDate)) {
        existingTask.startDate = incomingTask.startDate;
        changed = true;
    }
    if (incomingTask.completeDate && (isBadDate(existingTask.completeDate) || existingTask.completeDate !== incomingTask.completeDate)) {
        existingTask.completeDate = incomingTask.completeDate;
        changed = true;
    }
    if (incomingTask.details && existingTask.details !== incomingTask.details) {
        existingTask.details = incomingTask.details;
        changed = true;
    }
    if (sourceRow[columns.status] && existingTask.status !== incomingTask.status) {
        existingTask.status = incomingTask.status;
        changed = true;
    }
    if (sourceRow[columns.progress] !== undefined && sourceRow[columns.progress] !== "" && existingTask.progress !== incomingTask.progress) {
        existingTask.progress = incomingTask.progress;
        changed = true;
    }
    if (incomingTask.remarks && existingTask.remarks !== incomingTask.remarks) {
        existingTask.remarks = incomingTask.remarks;
        changed = true;
    }

    return changed;
}

function getPresenceStatusHTML() {
    return viewerPresence.online
        ? '<span style="color:var(--green);font-weight:700;">🟢 Online</span>'
        : '<span style="color:var(--red);font-weight:600;">🔴 Offline</span>';
}

function updateOwnPresence() {
    if (!loggedIn || !currentUser?.rc) return Promise.resolve();
    return db.ref(`ees_presence/${safeFirebaseKey(currentUser.rc)}`).set(firebase.database.ServerValue.TIMESTAMP);
}

function setupOwnPresence() {
    if (!currentUser?.rc) return;
    const presenceRef = db.ref(`ees_presence/${safeFirebaseKey(currentUser.rc)}`);
    presenceRef.set(firebase.database.ServerValue.TIMESTAMP);
    presenceRef.onDisconnect().set(0);
}

// ── Day 5: Search + pagination helpers ───────────────────────────
function normalizeSearchText(value) {
    return String(value ?? "").trim().toLowerCase();
}

function workOrderMatchesSearch(wo, query) {
    const q = normalizeSearchText(query);
    if (!q) return true;

    const assigneeNames = (Array.isArray(wo.assignees) ? wo.assignees : [])
        .map(rc => WORKERS.find(worker => worker.rc === rc)?.name || rc)
        .join(" ");

    const taskText = (Array.isArray(wo.tasks) ? wo.tasks : [])
        .map(task => `${task.taskNo || ""} ${task.details || ""} ${task.status || ""} ${task.remarks || ""}`)
        .join(" ");

    const searchable = [
        wo.id, wo.asset, wo.sr, wo.svo, wo.priority, wo.date, assigneeNames, taskText
    ].join(" ").toLowerCase();

    return searchable.includes(q);
}

function getOrderViewState(value = view) {
    if (!orderViewState[value]) {
        orderViewState[value] = { filterAsset: "All", search: "", page: 1, pageSize: 50 };
    }
    return orderViewState[value];
}

function resetOrderPaging(value = view) {
    getOrderViewState(value).page = 1;
}

function getPagedItems(items, state = getOrderViewState()) {
    const pageSize = clampNumber(state.pageSize, 10, 200);
    const totalItems = items.length;
    const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));
    const safePage = clampNumber(state.page, 1, totalPages);
    const start = (safePage - 1) * pageSize;

    state.page = safePage;
    state.pageSize = pageSize;

    return {
        items: items.slice(start, start + pageSize),
        totalItems,
        totalPages,
        page: safePage,
        pageSize,
        startIndex: totalItems === 0 ? 0 : start + 1,
        endIndex: Math.min(start + pageSize, totalItems)
    };
}

function isWorkOrderView(value = view) {
    return ["orders", "urgent", "ongoing", "completed"].includes(value);
}

function getWorkOrderViewConfig(value = view) {
    const activeOrders = orders.filter(order => order.overallProgress < 100);

    if (value === "completed") {
        return {
            isCompleted: true,
            emptyLabel: "completed",
            matchLabel: "completed",
            list: orders.filter(order => order.overallProgress === 100)
        };
    }

    if (value === "urgent") {
        return {
            isCompleted: false,
            emptyLabel: "urgent/critical",
            matchLabel: "urgent/critical active",
            list: activeOrders.filter(order => order.priority === "Urgent" || order.priority === "Critical")
        };
    }

    if (value === "ongoing") {
        return {
            isCompleted: false,
            emptyLabel: "ongoing",
            matchLabel: "ongoing active",
            list: activeOrders.filter(order => order.tasks.some(task => task.status === "Ongoing"))
        };
    }

    return {
        isCompleted: false,
        emptyLabel: "active",
        matchLabel: "active",
        list: activeOrders
    };
}

function renderPaginationControls(paged) {

    if (!paged || paged.totalItems <= paged.pageSize) return "";

    return `<div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:10px;margin:14px 0 4px;padding:12px;background:var(--card);border:1px solid var(--border);border-radius:10px;">
        <div style="font-size:12px;color:var(--muted);font-weight:600;">Showing ${paged.startIndex}-${paged.endIndex} of ${paged.totalItems}</div>
        <div style="display:flex;align-items:center;gap:8px;">
            <button class="btn btn-ghost" onclick="setOrderPage(${paged.page - 1})" ${paged.page <= 1 ? "disabled" : ""}>← Prev</button>
            <span style="font-size:12px;color:var(--text);font-weight:700;min-width:80px;text-align:center;">Page ${paged.page} / ${paged.totalPages}</span>
            <button class="btn btn-ghost" onclick="setOrderPage(${paged.page + 1})" ${paged.page >= paged.totalPages ? "disabled" : ""}>Next →</button>
        </div>
    </div>`;
}


// ── App State ─────────────────────────────────────────────────────
let loggedIn=false, currentUser=null, loginErr="", view="dashboard";
window.currentUser = currentUser;
window.currentView = view;
function publishCurrentUser() {
    window.currentUser = currentUser;
    window.dispatchEvent(new CustomEvent("ees:userchange", { detail: currentUser }));
}
function publishAppView() {
    window.currentView = view;
    window.dispatchEvent(new CustomEvent("ees:viewchange", { detail: { view } }));
}
let orders=[], workerLeaves={};
let selectedWO=null, selectedWorker=null, leaveModalWorker=null;
const orderViewState = {
    orders: { filterAsset: "All", search: "", page: 1, pageSize: 50 },
    urgent: { filterAsset: "All", search: "", page: 1, pageSize: 50 },
    ongoing: { filterAsset: "All", search: "", page: 1, pageSize: 50 },
    completed: { filterAsset: "All", search: "", page: 1, pageSize: 50 }
};
let isLightMode=false, isSidebarOpen=false, isReportsMenuOpen=false, isWOViewsMenuOpen=false, lastUpdate="Never";
let realtimeListeners = [];
let presencePingTimer = null;
let viewerPresence = { rc: "5794", name: "Ahmed Miushaan", online: false, lastSeen: 0 };

// ── Loading UI ────────────────────────────────────────────────────
function showLoad(msg="Loading...") {
    document.getElementById('loading-msg').textContent = msg;
    document.getElementById('loading-overlay').style.display = 'flex';
}
function hideLoad() {
    document.getElementById('loading-overlay').style.display = 'none';
}

// ── Toast ─────────────────────────────────────────────────────────
window.showToast = function(msg, isError=false) {
    const tc = document.getElementById('toast-container');
    tc.innerHTML = `<div class="toast${isError?' error':''}">${escapeHtml(msg)}</div>`;
    setTimeout(() => { tc.innerHTML=''; }, 3500);
};

// ── Firebase DB Functions ─────────────────────────────────────────

// Save a single work order + its tasks to Firebase
async function dbSaveWO(wo) {
    assertAdmin("save work orders");
    const woData = normalizeWorkOrder(wo);
    const key = safeFirebaseKey(woData.id);
    if (!key) throw new Error("Invalid work order ID. Cannot save.");
    await db.ref(`ees_wo/${key}`).set(woData);
}

// Save all changed WOs in one batch (used after Excel import)
async function dbBatchSave(changedWOs) {
    assertAdmin("import work orders");
    const updates = {};
    changedWOs.forEach(wo => {
        const woData = normalizeWorkOrder(wo);
        const key = safeFirebaseKey(woData.id);
        if (!key) throw new Error(`Invalid work order ID: ${woData.id}`);
        updates[`ees_wo/${key}`] = woData;
    });
    await db.ref().update(updates);
}

// Save leave for one worker
async function dbSaveLeave(rc, data) {
    assertAdmin("update leave status");
    const key = safeFirebaseKey(rc);
    if (!key) throw new Error("Invalid worker RC. Cannot save leave.");
    if(!data || data.type==='None') {
        await db.ref(`ees_leaves/${key}`).remove();
    } else {
        await db.ref(`ees_leaves/${key}`).set({
            type: String(data.type || "").trim(),
            from: formatForInput(data.from) || "",
            to: formatForInput(data.to) || ""
        });
    }
}

// Save last import timestamp
async function dbSaveLastUpdate(ts) {
    assertAdmin("update import metadata");
    await db.ref('ees_meta/lastUpdate').set(ts);
}

// Load all data from Firebase once
async function dbLoadAll() {
    const [woSnap, leavesSnap, metaSnap] = await Promise.all([
        db.ref('ees_wo').once('value'),
        db.ref('ees_leaves').once('value'),
        db.ref('ees_meta').once('value')
    ]);

    const woData = woSnap.val() || {};
    orders = Object.values(woData).map(wo => ({
        id: wo.id||"",
        asset: wo.asset||"",
        date: wo.date||"",
        sr: wo.sr||"",
        svo: wo.svo||"",
        priority: wo.priority||"Minor",
        overallProgress: wo.overallProgress||0,
        assignees: Array.isArray(wo.assignees) ? wo.assignees : [],
        tasks: Array.isArray(wo.tasks) ? wo.tasks : []
    }));

    const lvData = leavesSnap.val() || {};
    workerLeaves = {};
    Object.entries(lvData).forEach(([rc, lv]) => {
        if(lv && lv.type && lv.type !== 'None') workerLeaves[rc] = lv;
    });

    const meta = metaSnap.val() || {};
    lastUpdate = meta.lastUpdate || "Never";
}

// ── Real-time listener: auto-refresh when any device changes data ──
function startRealtimeSync() {
    stopRealtimeSync();

    const woRef = db.ref('ees_wo');
    const handler = woRef.on('value', snap => {
        if(!loggedIn) return;
        const data = snap.val() || {};
        orders = Object.values(data).map(wo => normalizeWorkOrder({
            id: wo.id||"", asset: wo.asset||"", date: wo.date||"",
            sr: wo.sr||"", svo: wo.svo||"",
            priority: wo.priority||"Minor",
            overallProgress: wo.overallProgress||0,
            assignees: Array.isArray(wo.assignees) ? wo.assignees : [],
            tasks: Array.isArray(wo.tasks) ? wo.tasks : []
        }));
        renderApp();
    });
    realtimeListeners.push({ ref: woRef, handler });

    const lvRef = db.ref('ees_leaves');
    const lvHandler = lvRef.on('value', snap => {
        if(!loggedIn) return;
        const data = snap.val() || {};
        workerLeaves = {};
        Object.entries(data).forEach(([rc, lv]) => {
            if(lv && lv.type && lv.type !== 'None') workerLeaves[rc] = lv;
        });
        renderApp();
    });
    realtimeListeners.push({ ref: lvRef, handler: lvHandler });

    // Day 4: presence is now listener-driven, not fetched during every render.
    if (isAdminUser()) {
        const viewerPresenceRef = db.ref(`ees_presence/${safeFirebaseKey(viewerPresence.rc)}`);
        const viewerPresenceHandler = viewerPresenceRef.on('value', snap => {
            if(!loggedIn) return;
            const lastSeen = Number(snap.val() || 0);
            const nextOnline = lastSeen > 0 && (Date.now() - lastSeen) < 300000;
            const changed = viewerPresence.online !== nextOnline || viewerPresence.lastSeen !== lastSeen;
            viewerPresence = { ...viewerPresence, online: nextOnline, lastSeen };
            if (changed) renderApp();
        });
        realtimeListeners.push({ ref: viewerPresenceRef, handler: viewerPresenceHandler });
    }

    setupOwnPresence();
    presencePingTimer = setInterval(() => {
        updateOwnPresence().catch(error => console.warn("Presence ping failed", error));
    }, 60000);
}

function stopRealtimeSync() {
    realtimeListeners.forEach(({ ref, handler }) => ref.off('value', handler));
    realtimeListeners = [];
    if (presencePingTimer) {
        clearInterval(presencePingTimer);
        presencePingTimer = null;
    }
    viewerPresence = { ...viewerPresence, online: false, lastSeen: 0 };
}

// ── Auth ──────────────────────────────────────────────────────────

window.handleLogin = async function() {
    const email = document.getElementById('login-u').value.trim().toLowerCase();
    const pass  = document.getElementById('login-p').value.trim();
    if(!email||!pass) { loginErr="Please enter email and password."; renderApp(); return; }

    showLoad("Signing in...");
    try {
        const cred = await auth.signInWithEmailAndPassword(email, pass);
        const profile = await loadUserProfile(cred.user || auth.currentUser);
        if(!profile) { await auth.signOut(); hideLoad(); loginErr="Access denied for this account."; renderApp(); return; }

        currentUser = profile;
        loggedIn = true;
        publishCurrentUser();
        loginErr = "";
        view = "dashboard";

        showLoad("Loading data...");
        await dbLoadAll();
        await startUserSession();
        hideLoad();
        renderApp();
        startRealtimeSync();
    } catch(e) {
        hideLoad();
        if(e.code === 'auth/wrong-password' || e.code === 'auth/user-not-found' || e.code === 'auth/invalid-credential') {
            loginErr = "Wrong email or password.";
        } else if(e.code === 'auth/invalid-api-key' || e.code === 'auth/configuration-not-found') {
            loginErr = "Firebase not configured yet. Paste your firebaseConfig above.";
        } else if(e.message && e.message.includes('fetch')) {
            loginErr = "Cannot connect. Check your Firebase config in the code.";
        } else {
            loginErr = "Error: " + e.message;
        }
        renderApp();
    }
};

window.handleLogout = async function() {
    stopRealtimeSync();
    await endUserSession();
    await auth.signOut();
    loggedIn = false; currentUser = null; view = "dashboard"; publishCurrentUser(); publishAppView(); orders = []; workerLeaves = {};
    renderApp();
};

// ── UI Handlers ───────────────────────────────────────────────────

window.toggleTheme = () => { isLightMode=!isLightMode; document.documentElement[isLightMode?'setAttribute':'removeAttribute']('data-theme','light'); renderApp(); };
window.toggleSidebar = () => { isSidebarOpen=!isSidebarOpen; isReportsMenuOpen=false; isWOViewsMenuOpen=false; renderApp(); };
window.toggleReportsMenu = function(event) {
    if (event) event.stopPropagation();
    isWOViewsMenuOpen = false;
    isReportsMenuOpen = !isReportsMenuOpen;
    renderApp();
};
window.toggleWOViewsMenu = function(event) {
    if (event) event.stopPropagation();
    isReportsMenuOpen = false;
    isWOViewsMenuOpen = !isWOViewsMenuOpen;
    renderApp();
};
window.runReportAction = function(action) {
    isReportsMenuOpen = false;
    isWOViewsMenuOpen = false;
    if (action === "excel") return exportExcel();
    if (action === "quickPdf") return exportQuickPDF();
    if (action === "fullPdf") return exportFullPDF();
    if (action === "jsonText") return exportJsonText();
    showToast("⚠️ Unknown export option", true);
};
window.setView = v => {
    if (v === "upload" && !requireAdmin("manage data imports")) return;
    view=v; selectedWO=null; selectedWorker=null; leaveModalWorker=null; isSidebarOpen=false; isReportsMenuOpen=false; isWOViewsMenuOpen=false; renderApp(); publishAppView();
};
window.setFilter = f => {
    const state = getOrderViewState("orders");
    state.filterAsset = f || "All";
    resetOrderPaging("orders");
    renderApp();
};
window.applyOrderSearch = () => {
    const state = getOrderViewState("orders");
    const searchInput = document.getElementById("active-order-search");
    state.search = String(searchInput ? searchInput.value : state.search || "");
    resetOrderPaging("orders");
    renderApp();
};

window.setOrderSearch = q => {
    const state = getOrderViewState("orders");
    state.search = String(q || "");
    resetOrderPaging("orders");
    renderApp();
};
window.setOrderPage = () => {};
window.setOrderPageSize = () => {};
window.selectOrder = id => { selectedWO=id; renderApp(); };
window.selectWorker = rc => { selectedWorker=rc; renderApp(); };
window.closeWorkerModal = () => { selectedWorker=null; renderApp(); };
window.openLeaveModal = rc => {
    if (!requireAdmin("edit leave status")) return;
    leaveModalWorker=rc; renderApp();
};
window.closeLeaveModal = () => { leaveModalWorker=null; renderApp(); };

window.togglePassword = function() {
    const p=document.getElementById('login-p'), e=document.getElementById('eye-icon');
    e.style.transform='scale(.7)';
    setTimeout(()=>{
        if(p.type==='password'){p.type='text';e.style.stroke='var(--blue)';e.innerHTML=`<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle>`;}
        else{p.type='password';e.style.stroke='var(--muted)';e.innerHTML=`<path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"></path><line x1="1" y1="1" x2="23" y2="23"></line>`;}
        e.style.transform='scale(1)';
    },150);
};

window.saveLeaveUpdates = async function(rc) {
    if (!requireAdmin("update leave status")) return;
    const type = document.getElementById('l-type').value;
    const from = document.getElementById('l-from')?.value||'';
    const to   = document.getElementById('l-to')?.value||'';
    if(type==='None') delete workerLeaves[rc];
    else workerLeaves[rc] = {type,from,to};
    closeLeaveModal(); renderApp();
    try {
        await dbSaveLeave(rc, type==='None'?null:{type,from,to});
        await auditLog("LEAVE_UPDATED", { rc, type, from, to });
        showToast("✅ Leave status saved");
    }
    catch(e) { showToast("❌ Failed to save leave", true); }
};

window.openWOFromWorker = function(id) {
    selectedWorker=null;
    const wo=orders.find(o=>o.id===id);
    view=(wo&&wo.overallProgress===100)?"completed":"orders";
    selectedWO=id;
    resetOrderPaging();
    renderApp();
};

window.clearDatabase = async function() {
    if (!requireAdmin("clear the database")) return;

    const typed = prompt("🚨 DANGER: This will BACKUP then DELETE active WO, leave, and meta data. Type CLEAR EES to continue.");
    if (typed !== "CLEAR EES") {
        showToast("Database clear cancelled.");
        return;
    }

    showLoad("Creating safety backup...");
    let backupPath = "";

    try {
        backupPath = await backupOperationalData("clear_database_before_delete");
        await auditLog("DATABASE_CLEAR_BACKUP_CREATED", { backupPath });

        showLoad("Clearing database...");
        await Promise.all([
            db.ref('ees_wo').remove(),
            db.ref('ees_leaves').remove(),
            db.ref('ees_meta').remove()
        ]);

        orders=[]; workerLeaves={}; lastUpdate="Never";
        await auditLog("DATABASE_CLEARED", { backupPath });
        showToast("🗑️ Database cleared after backup");
    } catch(e) {
        console.error(e);
        showToast("❌ Safe clear failed: " + e.message, true);
    }

    hideLoad(); renderApp();
};

// ── WO Modal sync ─────────────────────────────────────────────────

function syncModal() {
    if(!selectedWO) return;
    const wo = orders.find(o=>o.id===selectedWO);
    if(!wo) return;

    const priorityInput = document.getElementById('wo-priority');
    if (priorityInput) wo.priority = priorityInput.value;

    wo.tasks.forEach((t,i) => {
        const statusInput = document.getElementById(`t-stat-${i}`);
        const progressInput = document.getElementById(`t-prog-${i}`);
        const startInput = document.getElementById(`t-start-${i}`);
        const endInput = document.getElementById(`t-end-${i}`);
        const remarksInput = document.getElementById(`t-rem-${i}`);

        if (statusInput) t.status = statusInput.value || t.status;

        // Preserve the old progress if the field is accidentally left blank.
        if (progressInput && progressInput.value.trim() !== '') {
            t.progress = clampNumber(progressInput.value, 0, 100);
        }

        t.startDate = startInput?.value || '';
        t.completeDate = endInput?.value || '';
        t.remarks = remarksInput?.value || '';
    });
}

window.addAssignee = function(rc) {
    if (!requireAdmin("assign employees")) return;
    if(!rc) return;
    syncModal();
    const wo=orders.find(o=>o.id===selectedWO);
    if(!wo) { showToast("⚠️ Work order no longer exists.", true); selectedWO=null; renderApp(); return; }
    if(!wo.assignees.includes(rc)) wo.assignees.push(rc);
    renderApp();
};
window.removeAssignee = function(rc) {
    if (!requireAdmin("remove assignees")) return;
    syncModal();
    const wo=orders.find(o=>o.id===selectedWO);
    if(!wo) { showToast("⚠️ Work order no longer exists.", true); selectedWO=null; renderApp(); return; }
    wo.assignees=wo.assignees.filter(x=>x!==rc);
    renderApp();
};

window.saveWOUpdates = async function(woId) {
    if (!requireAdmin("save work orders")) return;
    syncModal();
    const wo = orders.find(o=>o.id===woId);

    if(!wo) {
        selectedWO = null;
        showToast("⚠️ This work order no longer exists. View refreshed.", true);
        renderApp();
        return;
    }

    let tp=0;
    wo.tasks.forEach(t => {
        t.progress = clampNumber(t.progress, 0, 100);
        tp += t.progress;
    });
    wo.overallProgress = wo.tasks.length>0 ? Math.round(tp/wo.tasks.length) : 0;
    renderApp(); // instant UI
    try {
        await dbSaveWO(wo);
        await auditLog("WO_UPDATED", { workOrderId: wo.id, progress: wo.overallProgress, taskCount: wo.tasks.length });
        showToast(`✅ Saved ${wo.id}`);
    }
    catch(e) { console.error(e); showToast("❌ Save failed — check connection", true); }
};

// ── Excel Export ──────────────────────────────────────────────────

window.exportExcel = function() {
    if(!orders.length){showToast("No data to export.");return;}
    if(!window.XLSX){showToast("⏳ Library loading, try again...",true);return;}
    showToast("📊 Building multi-sheet report...");

    try {
        const wb = XLSX.utils.book_new();
        const today = new Date().toLocaleDateString("en-GB", {day:"numeric", month:"long", year:"numeric"});
        const ts = new Date().toLocaleString("en-US", {dateStyle:"medium", timeStyle:"short"});
        const userName = (currentUser && currentUser.name) || "System";

        // Style color palette
        const COL = {
            header:       { fgColor:{rgb:"0A84FF"}, color:{rgb:"FFFFFF"}, bold:true },
            titleBig:     { fgColor:{rgb:"1C1C1E"}, color:{rgb:"FFFFFF"}, bold:true, size:16 },
            critical:     { fgColor:{rgb:"FF3B30"}, color:{rgb:"FFFFFF"}, bold:true },
            urgent:       { fgColor:{rgb:"FF9500"}, color:{rgb:"FFFFFF"}, bold:true },
            major:        { fgColor:{rgb:"FFCC00"}, color:{rgb:"000000"}, bold:true },
            minor:        { fgColor:{rgb:"8E8E93"}, color:{rgb:"FFFFFF"}, bold:true },
            ongoing:      { fgColor:{rgb:"CCE5FF"}, color:{rgb:"003380"}, bold:true },
            completed:    { fgColor:{rgb:"D4F4DD"}, color:{rgb:"14532D"}, bold:true },
            pending:      { fgColor:{rgb:"E5E5EA"}, color:{rgb:"3C3C43"}, bold:true },
            onhold:       { fgColor:{rgb:"FFE8CC"}, color:{rgb:"7C2D12"}, bold:true },
            cancelled:    { fgColor:{rgb:"FFCCCC"}, color:{rgb:"7F1D1D"}, bold:true },
            sectionHead:  { fgColor:{rgb:"E5F0FF"}, color:{rgb:"003380"}, bold:true, size:12 },
            total:        { fgColor:{rgb:"1C1C1E"}, color:{rgb:"FFCC00"}, bold:true, size:14 },
            label:        { color:{rgb:"6B7280"}, bold:false },
            value:        { color:{rgb:"000000"}, bold:true },
            leaveRow:     { fgColor:{rgb:"FFE4E6"}, color:{rgb:"7F1D1D"} }
        };

        const priColor = p => p==="Critical"?COL.critical : p==="Urgent"?COL.urgent : p==="Major"?COL.major : COL.minor;
        const statColor = s => s==="Ongoing"?COL.ongoing : s==="Completed"?COL.completed : s==="Onhold"?COL.onhold : s==="Cancelled"?COL.cancelled : COL.pending;

        // Helper: apply cell style
        const setStyle = (ws, addr, fill) => {
            if (!ws[addr]) ws[addr] = { v: "", t: "s" };
            ws[addr].s = {
                fill: fill.fgColor ? { patternType:"solid", fgColor:fill.fgColor } : undefined,
                font: { color: fill.color, bold: fill.bold, sz: fill.size || 11, name:"Calibri" },
                alignment: { vertical:"center", horizontal:"left", wrapText:true },
                border: { top:{style:"thin",color:{rgb:"E5E7EB"}}, bottom:{style:"thin",color:{rgb:"E5E7EB"}}, left:{style:"thin",color:{rgb:"E5E7EB"}}, right:{style:"thin",color:{rgb:"E5E7EB"}} }
            };
        };

        const setCell = (ws, addr, val, fill) => {
            ws[addr] = { v: val, t: typeof val==="number" ? "n" : "s" };
            if (fill) setStyle(ws, addr, fill);
        };

        const colLetter = (n) => { let s=""; while(n>=0){s=String.fromCharCode(65+(n%26))+s;n=Math.floor(n/26)-1;} return s; };

        // ═══════════════════════════════════════════════════════════
        // SHEET 1: SUMMARY
        // ═══════════════════════════════════════════════════════════
        {
            const ws = {};
            const merges = [];
            let r = 0;

            // Big title
            setCell(ws, `A${r+1}`, "EES WO CONTROL DATABASE", COL.titleBig);
            merges.push({s:{r:r,c:0},e:{r:r,c:3}});
            r++;
            setCell(ws, `A${r+1}`, "MTCC · Electrical & Electronic Services", COL.label);
            merges.push({s:{r:r,c:0},e:{r:r,c:3}});
            r += 2;

            setCell(ws, `A${r+1}`, "Report Date:", COL.label);
            setCell(ws, `B${r+1}`, today, COL.value);
            r++;
            setCell(ws, `A${r+1}`, "Generated By:", COL.label);
            setCell(ws, `B${r+1}`, userName, COL.value);
            r++;
            setCell(ws, `A${r+1}`, "Generated At:", COL.label);
            setCell(ws, `B${r+1}`, ts, COL.value);
            r += 2;

            // Work Order Stats
            setCell(ws, `A${r+1}`, "WORK ORDER STATISTICS", COL.sectionHead);
            merges.push({s:{r:r,c:0},e:{r:r,c:3}});
            r++;

            const total = orders.length;
            const completed = orders.filter(o => o.overallProgress === 100).length;
            const active = total - completed;
            const completionRate = total > 0 ? Math.round((completed/total)*100) : 0;

            setCell(ws, `A${r+1}`, "Total Work Orders", COL.label); setCell(ws, `B${r+1}`, total, COL.total); r++;
            setCell(ws, `A${r+1}`, "Active WOs", COL.label); setCell(ws, `B${r+1}`, active, COL.value); r++;
            setCell(ws, `A${r+1}`, "Completed WOs", COL.label); setCell(ws, `B${r+1}`, completed, COL.value); r++;
            setCell(ws, `A${r+1}`, "Completion Rate", COL.label); setCell(ws, `B${r+1}`, completionRate + "%", COL.value); r++;
            r++;

            // Priority Breakdown
            setCell(ws, `A${r+1}`, "PRIORITY BREAKDOWN (Active Only)", COL.sectionHead);
            merges.push({s:{r:r,c:0},e:{r:r,c:3}});
            r++;

            const activeOrders = orders.filter(o => o.overallProgress < 100);
            const priCounts = {Critical:0, Urgent:0, Major:0, Minor:0};
            activeOrders.forEach(o => { if(priCounts[o.priority] !== undefined) priCounts[o.priority]++; });

            ["Critical","Urgent","Major","Minor"].forEach(p => {
                setCell(ws, `A${r+1}`, p, priColor(p));
                setCell(ws, `B${r+1}`, priCounts[p], COL.value);
                r++;
            });
            r++;

            // Task Status
            setCell(ws, `A${r+1}`, "TASK STATUS BREAKDOWN", COL.sectionHead);
            merges.push({s:{r:r,c:0},e:{r:r,c:3}});
            r++;

            let sOng=0,sOh=0,sP=0,sC=0,sCn=0,sT=0;
            orders.forEach(o => o.tasks.forEach(t => {
                sT++;
                if(t.status==="Ongoing") sOng++;
                else if(t.status==="Onhold") sOh++;
                else if(t.status==="Pending") sP++;
                else if(t.status==="Completed") sC++;
                else if(t.status==="Cancelled") sCn++;
            }));

            setCell(ws, `A${r+1}`, "Total Tasks", COL.label); setCell(ws, `B${r+1}`, sT, COL.total); r++;
            setCell(ws, `A${r+1}`, "Ongoing", COL.ongoing); setCell(ws, `B${r+1}`, sOng, COL.value); r++;
            setCell(ws, `A${r+1}`, "Onhold", COL.onhold); setCell(ws, `B${r+1}`, sOh, COL.value); r++;
            setCell(ws, `A${r+1}`, "Pending", COL.pending); setCell(ws, `B${r+1}`, sP, COL.value); r++;
            setCell(ws, `A${r+1}`, "Completed", COL.completed); setCell(ws, `B${r+1}`, sC, COL.value); r++;
            setCell(ws, `A${r+1}`, "Cancelled", COL.cancelled); setCell(ws, `B${r+1}`, sCn, COL.value); r++;
            r++;

            // Team Status
            setCell(ws, `A${r+1}`, "TEAM STATUS", COL.sectionHead);
            merges.push({s:{r:r,c:0},e:{r:r,c:3}});
            r++;

            let onLeave = 0;
            WORKERS.forEach(w => { if(workerLeaves[w.rc]?.type && workerLeaves[w.rc].type !== "None") onLeave++; });

            setCell(ws, `A${r+1}`, "Total Headcount", COL.label); setCell(ws, `B${r+1}`, WORKERS.length, COL.total); r++;
            setCell(ws, `A${r+1}`, "On Duty", COL.ongoing); setCell(ws, `B${r+1}`, WORKERS.length - onLeave, COL.value); r++;
            setCell(ws, `A${r+1}`, "On Leave", COL.cancelled); setCell(ws, `B${r+1}`, onLeave, COL.value); r++;

            ws["!ref"] = `A1:D${r+2}`;
            ws["!cols"] = [{wch:28},{wch:22},{wch:16},{wch:16}];
            ws["!merges"] = merges;
            XLSX.utils.book_append_sheet(wb, ws, "Summary");
        }

        // ═══════════════════════════════════════════════════════════
        // Helper to build WO sheet
        // ═══════════════════════════════════════════════════════════
        const buildWOSheet = (sheetName, woList) => {
            const ws = {};
            const headers = ["WO Number","Asset","Date","Priority","Assignees","Total Tasks","Ongoing","Completed","Progress %","Status"];

            // Header row
            headers.forEach((h, i) => setCell(ws, `${colLetter(i)}1`, h, COL.header));

            woList.forEach((wo, idx) => {
                const r = idx + 2;
                const names = wo.assignees.map(rc => WORKERS.find(w=>w.rc===rc)?.name || rc).join(", ") || "—";
                const onCount = wo.tasks.filter(t => t.status==="Ongoing").length;
                const compCount = wo.tasks.filter(t => t.status==="Completed").length;
                const overallStatus = wo.overallProgress === 100 ? "Completed" : onCount > 0 ? "Ongoing" : "Pending";

                setCell(ws, `A${r}`, wo.id, COL.value);
                setCell(ws, `B${r}`, wo.asset || "—");
                setCell(ws, `C${r}`, formatDateNice(wo.date));
                setCell(ws, `D${r}`, wo.priority, priColor(wo.priority));
                setCell(ws, `E${r}`, names);
                setCell(ws, `F${r}`, wo.tasks.length);
                setCell(ws, `G${r}`, onCount);
                setCell(ws, `H${r}`, compCount);
                setCell(ws, `I${r}`, wo.overallProgress + "%", wo.overallProgress===100?COL.completed:wo.overallProgress>0?COL.ongoing:COL.pending);
                setCell(ws, `J${r}`, overallStatus, statColor(overallStatus));
            });

            ws["!ref"] = `A1:${colLetter(headers.length-1)}${woList.length+1}`;
            ws["!cols"] = [{wch:20},{wch:28},{wch:16},{wch:11},{wch:30},{wch:12},{wch:10},{wch:11},{wch:12},{wch:14}];
            ws["!freeze"] = { xSplit: 0, ySplit: 1 };
            ws["!autofilter"] = { ref: `A1:${colLetter(headers.length-1)}${woList.length+1}` };
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
        };

        // SHEET 2: Active WOs
        const activeList = orders.filter(o => o.overallProgress < 100)
            .sort((a,b) => {
                const order = {Critical:0, Urgent:1, Major:2, Minor:3};
                return (order[a.priority]||9) - (order[b.priority]||9);
            });
        buildWOSheet("Active WOs", activeList);

        // SHEET 3: Completed WOs
        const completedList = orders.filter(o => o.overallProgress === 100);
        buildWOSheet("Completed WOs", completedList);

        // ═══════════════════════════════════════════════════════════
        // SHEET 4: All Tasks (granular)
        // ═══════════════════════════════════════════════════════════
        {
            const ws = {};
            const headers = ["WO Number","Asset","Priority","Task No","Task Details","Status","Start Date","Progress %","Complete Date","Remarks","Assignees"];
            headers.forEach((h, i) => setCell(ws, `${colLetter(i)}1`, h, COL.header));

            let r = 2;
            orders.forEach(wo => {
                const names = wo.assignees.map(rc => WORKERS.find(w=>w.rc===rc)?.name || rc).join(", ") || "—";
                wo.tasks.forEach(t => {
                    setCell(ws, `A${r}`, wo.id, COL.value);
                    setCell(ws, `B${r}`, wo.asset || "—");
                    setCell(ws, `C${r}`, wo.priority, priColor(wo.priority));
                    setCell(ws, `D${r}`, t.taskNo);
                    setCell(ws, `E${r}`, t.details || "");
                    setCell(ws, `F${r}`, t.status, statColor(t.status));
                    setCell(ws, `G${r}`, formatDateNice(t.startDate));
                    setCell(ws, `H${r}`, (t.progress||0) + "%");
                    setCell(ws, `I${r}`, formatDateNice(t.completeDate));
                    setCell(ws, `J${r}`, t.remarks || "");
                    setCell(ws, `K${r}`, names);
                    r++;
                });
            });

            ws["!ref"] = `A1:K${r}`;
            ws["!cols"] = [{wch:20},{wch:26},{wch:11},{wch:11},{wch:40},{wch:13},{wch:14},{wch:11},{wch:14},{wch:28},{wch:28}];
            ws["!freeze"] = { xSplit: 0, ySplit: 1 };
            ws["!autofilter"] = { ref: `A1:K${r-1}` };
            XLSX.utils.book_append_sheet(wb, ws, "All Tasks");
        }

        // ═══════════════════════════════════════════════════════════
        // SHEET 5: Team Workload
        // ═══════════════════════════════════════════════════════════
        {
            const ws = {};
            const headers = ["RC","Worker Name","Designation","Total WOs","Ongoing","Completed","Leave Status","From","To"];
            headers.forEach((h, i) => setCell(ws, `${colLetter(i)}1`, h, COL.header));

            WORKERS.forEach((w, idx) => {
                const r = idx + 2;
                const woList = orders.filter(o => o.assignees && o.assignees.includes(w.rc));
                const ongoing = woList.filter(o => o.tasks.some(t => t.status==="Ongoing")).length;
                const completed = woList.filter(o => o.overallProgress===100).length;
                const lv = workerLeaves[w.rc];
                const leaveStatus = lv && lv.type !== "None" ? lv.type : "Active";

                setCell(ws, `A${r}`, w.rc, COL.value);
                setCell(ws, `B${r}`, w.name, COL.value);
                setCell(ws, `C${r}`, w.designation);
                setCell(ws, `D${r}`, woList.length);
                setCell(ws, `E${r}`, ongoing, ongoing>0?COL.ongoing:undefined);
                setCell(ws, `F${r}`, completed, completed>0?COL.completed:undefined);
                setCell(ws, `G${r}`, leaveStatus, leaveStatus!=="Active"?COL.cancelled:COL.completed);
                setCell(ws, `H${r}`, lv?.from ? formatDateNice(lv.from) : "—");
                setCell(ws, `I${r}`, lv?.to ? formatDateNice(lv.to) : "—");
            });

            ws["!ref"] = `A1:I${WORKERS.length+1}`;
            ws["!cols"] = [{wch:8},{wch:28},{wch:26},{wch:11},{wch:11},{wch:12},{wch:16},{wch:14},{wch:14}];
            ws["!freeze"] = { xSplit: 0, ySplit: 1 };
            ws["!autofilter"] = { ref: `A1:I${WORKERS.length+1}` };
            XLSX.utils.book_append_sheet(wb, ws, "Team Workload");
        }

        // ═══════════════════════════════════════════════════════════
        // SHEET 6: Leave Register
        // ═══════════════════════════════════════════════════════════
        {
            const ws = {};
            const headers = ["RC","Worker Name","Designation","Leave Type","From","To","Days"];
            headers.forEach((h, i) => setCell(ws, `${colLetter(i)}1`, h, COL.header));

            const onLeaveWorkers = WORKERS.filter(w => workerLeaves[w.rc] && workerLeaves[w.rc].type !== "None");
            if(onLeaveWorkers.length === 0) {
                setCell(ws, `A2`, "No workers currently on leave", COL.label);
                ws["!ref"] = `A1:G2`;
            } else {
                onLeaveWorkers.forEach((w, idx) => {
                    const r = idx + 2;
                    const lv = workerLeaves[w.rc];
                    let days = "—";
                    if(lv.from && lv.to) {
                        const d1 = new Date(lv.from), d2 = new Date(lv.to);
                        if(!isNaN(d1) && !isNaN(d2)) days = Math.round((d2-d1)/(1000*60*60*24)) + 1;
                    }
                    setCell(ws, `A${r}`, w.rc, COL.value);
                    setCell(ws, `B${r}`, w.name, COL.leaveRow);
                    setCell(ws, `C${r}`, w.designation);
                    setCell(ws, `D${r}`, lv.type, COL.cancelled);
                    setCell(ws, `E${r}`, lv.from ? formatDateNice(lv.from) : "—");
                    setCell(ws, `F${r}`, lv.to ? formatDateNice(lv.to) : "—");
                    setCell(ws, `G${r}`, days);
                });
                ws["!ref"] = `A1:G${onLeaveWorkers.length+1}`;
            }

            ws["!cols"] = [{wch:8},{wch:28},{wch:26},{wch:14},{wch:14},{wch:14},{wch:8}];
            ws["!freeze"] = { xSplit: 0, ySplit: 1 };
            XLSX.utils.book_append_sheet(wb, ws, "Leave Register");
        }

        // ═══════════════════════════════════════════════════════════
        // Generate file
        // ═══════════════════════════════════════════════════════════
        const wbout = XLSX.write(wb, { bookType:"xlsx", type:"array", cellStyles:true });
        const blob = new Blob([wbout], { type:"application/octet-stream" });
        const url = URL.createObjectURL(blob);
        const fn = `EES_WO_Report_${new Date().toISOString().slice(0,10)}.xlsx`;

        const dlModal = `<div class="overlay" id="dlm" onclick="if(event.target===this)this.remove()">
            <div class="modal" style="max-width:420px;background:var(--surface);">
                <div class="modal-head"><div style="font-family:'Rajdhani',sans-serif;font-size:20px;font-weight:700;">Multi-Sheet Report Ready 📊</div><button onclick="document.getElementById('dlm').remove()" style="background:none;border:none;font-size:22px;color:var(--muted);cursor:pointer;">✕</button></div>
                <div class="modal-body" style="padding:25px;">
                    <div style="color:var(--muted);font-size:12px;margin-bottom:15px;line-height:1.6;">
                        <div style="color:var(--text);font-weight:600;margin-bottom:10px;">6 Sheets included:</div>
                        <div style="padding:4px 0;">📄 Summary — Stats & KPIs</div>
                        <div style="padding:4px 0;">🔵 Active WOs — ${orders.filter(o=>o.overallProgress<100).length} orders</div>
                        <div style="padding:4px 0;">🟢 Completed WOs — ${orders.filter(o=>o.overallProgress===100).length} orders</div>
                        <div style="padding:4px 0;">📋 All Tasks — granular</div>
                        <div style="padding:4px 0;">👥 Team Workload — ${WORKERS.length} workers</div>
                        <div style="padding:4px 0;">🏖️ Leave Register</div>
                    </div>
                    <a href="${url}" download="${fn}" class="btn btn-primary" style="text-decoration:none;padding:14px;width:100%;justify-content:center;font-size:14px;" onclick="setTimeout(()=>document.getElementById('dlm').remove(),500)">⬇️ Download Excel</a>
                    <button class="btn btn-ghost" style="margin-top:10px;width:100%;justify-content:center;" onclick="document.getElementById('dlm').remove()">Cancel</button>
                </div>
            </div>
        </div>`;
        document.body.insertAdjacentHTML('beforeend', dlModal);
        showToast("✅ Report generated!");

    } catch(err) {
        console.error(err);
        showToast("❌ Export failed: " + err.message, true);
    }
};

// ── PDF Report Builder (jsPDF + autoTable) ──────────────────────

function _priColor(p) {
    // Returns RGB tuple for priority colouring in PDF tables
    if (p === "Critical") return [255, 59, 48];
    if (p === "Urgent")   return [255, 149, 0];
    if (p === "Major")    return [255, 204, 0];
    return [142, 142, 147]; // Minor
}

function _statColor(s) {
    if (s === "Ongoing")   return [10, 132, 255];
    if (s === "Completed") return [50, 215, 75];
    if (s === "Onhold")    return [255, 149, 0];
    if (s === "Cancelled") return [255, 59, 48];
    return [142, 142, 147]; // Pending
}

function _pdfCheckJsPDF() {
    if (typeof window.jspdf === "undefined" || !window.jspdf.jsPDF) {
        showToast("⏳ PDF library loading, try again in a second...", true);
        return null;
    }
    return window.jspdf.jsPDF;
}

function _pdfAddHeader(doc, subtitle) {
    // Top letterhead band
    doc.setFillColor(10, 132, 255);
    doc.rect(0, 0, 210, 22, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.text("EES WO CONTROL DATABASE", 14, 10);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.text("MTCC · Electrical & Electronic Services", 14, 16);

    // Right-side subtitle
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.text(subtitle, 196, 13, { align: "right" });

    // Reset for body
    doc.setTextColor(0, 0, 0);
}

function _pdfAddMeta(doc, startY) {
    const today = new Date().toLocaleDateString("en-GB", { day:"numeric", month:"long", year:"numeric" });
    const nowTime = new Date().toLocaleTimeString("en-US", { hour:"2-digit", minute:"2-digit" });
    const userName = (currentUser && currentUser.name) || "System";
    const userRole = (currentUser && currentUser.designation) || "";

    doc.setFontSize(9);
    doc.setTextColor(75, 85, 99);
    doc.text(`Generated: ${today} at ${nowTime}`, 14, startY);
    doc.text(`Prepared By: ${userName}${userRole ? " · " + userRole : ""}`, 14, startY + 5);

    // Divider
    doc.setDrawColor(229, 231, 235);
    doc.setLineWidth(0.3);
    doc.line(14, startY + 8, 196, startY + 8);
    doc.setTextColor(0, 0, 0);
    return startY + 13;
}

function _pdfAddFooter(doc) {
    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(107, 114, 128);
        doc.setDrawColor(229, 231, 235);
        doc.line(14, 286, 196, 286);
        doc.text(`EES WO Control Database · Confidential`, 14, 291);
        doc.text(`Page ${i} of ${pageCount}`, 196, 291, { align: "right" });
    }
}

function _pdfKPIBoxes(doc, startY, stats) {
    // 3 boxes across the page
    const boxes = [
        { label: "ACTIVE WOs",    value: stats.active,     color: [10, 132, 255] },
        { label: "COMPLETED WOs", value: stats.completed,  color: [50, 215, 75] },
        { label: "ON DUTY",       value: `${stats.onDuty}/${stats.headcount}`, color: [107, 114, 128] }
    ];
    const boxW = 58, gap = 5, startX = 14;

    boxes.forEach((box, i) => {
        const x = startX + i * (boxW + gap);
        // Colored top band
        doc.setFillColor(...box.color);
        doc.rect(x, startY, boxW, 3, "F");
        // White card
        doc.setFillColor(249, 250, 251);
        doc.setDrawColor(229, 231, 235);
        doc.rect(x, startY + 3, boxW, 22, "FD");
        // Label
        doc.setTextColor(107, 114, 128);
        doc.setFont("helvetica", "bold");
        doc.setFontSize(8);
        doc.text(box.label, x + 4, startY + 9);
        // Value
        doc.setTextColor(...box.color);
        doc.setFontSize(18);
        doc.text(String(box.value), x + 4, startY + 20);
    });
    return startY + 32;
}

function _pdfBuildStats() {
    const total = orders.length;
    const completed = orders.filter(o => o.overallProgress === 100).length;
    const active = total - completed;
    let onLeave = 0;
    WORKERS.forEach(w => { if(workerLeaves[w.rc]?.type && workerLeaves[w.rc].type !== "None") onLeave++; });

    let tOn=0, tOh=0, tP=0, tC=0, tCn=0;
    orders.forEach(o => o.tasks.forEach(t => {
        if(t.status==="Ongoing") tOn++;
        else if(t.status==="Onhold") tOh++;
        else if(t.status==="Pending") tP++;
        else if(t.status==="Completed") tC++;
        else if(t.status==="Cancelled") tCn++;
    }));

    return {
        total, active, completed,
        headcount: WORKERS.length,
        onDuty: WORKERS.length - onLeave,
        onLeave,
        tOn, tOh, tP, tC, tCn
    };
}

function _pdfWOTable(doc, startY, title, woList, titleColor) {
    // Section title
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(...titleColor);
    doc.text(`${title}  (${woList.length})`, 14, startY);
    doc.setTextColor(0, 0, 0);

    if (woList.length === 0) {
        doc.setFont("helvetica", "italic");
        doc.setFontSize(9);
        doc.setTextColor(156, 163, 175);
        doc.text("No items in this section.", 14, startY + 6);
        doc.setTextColor(0, 0, 0);
        return startY + 12;
    }

    const head = [["WO Number", "Asset", "Date", "Priority", "Tasks", "Progress"]];
    const body = woList.map(wo => [
        wo.id,
        wo.asset || "—",
        formatDateNice(wo.date),
        wo.priority,
        String(wo.tasks.length),
        (wo.overallProgress || 0) + "%"
    ]);

    doc.autoTable({
        head, body,
        startY: startY + 3,
        theme: "grid",
        styles: { fontSize: 8, cellPadding: 2, textColor: [31, 41, 55], lineColor: [229, 231, 235] },
        headStyles: { fillColor: [31, 41, 55], textColor: 255, fontStyle: "bold", halign: "left" },
        alternateRowStyles: { fillColor: [249, 250, 251] },
        columnStyles: {
            0: { cellWidth: 36, fontStyle: "bold" },
            1: { cellWidth: 56 },
            2: { cellWidth: 28 },
            3: { cellWidth: 22, halign: "center" },
            4: { cellWidth: 16, halign: "center" },
            5: { cellWidth: 24, halign: "right", fontStyle: "bold" }
        },
        didParseCell: (data) => {
            // Color priority cell
            if (data.section === "body" && data.column.index === 3) {
                const [r, g, b] = _priColor(data.cell.raw);
                data.cell.styles.fillColor = [r, g, b];
                data.cell.styles.textColor = 255;
                data.cell.styles.fontStyle = "bold";
            }
            // Progress column colour
            if (data.section === "body" && data.column.index === 5) {
                const prog = parseInt(data.cell.raw) || 0;
                if (prog === 100) data.cell.styles.textColor = [22, 163, 74];
                else if (prog > 0) data.cell.styles.textColor = [10, 132, 255];
                else data.cell.styles.textColor = [107, 114, 128];
            }
        },
        margin: { left: 14, right: 14 }
    });

    return doc.lastAutoTable.finalY + 8;
}

function _pdfTeamSection(doc, startY) {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(10, 132, 255);
    doc.text(`👥 TEAM WORKLOAD  (${WORKERS.length})`, 14, startY);
    doc.setTextColor(0, 0, 0);

    const head = [["Worker", "Designation", "RC", "Active WOs", "Status"]];
    const body = WORKERS.map(w => {
        const wos = orders.filter(o => o.assignees && o.assignees.includes(w.rc));
        const activeCount = wos.filter(o => o.overallProgress < 100).length;
        const lv = workerLeaves[w.rc];
        const leaveStatus = lv && lv.type !== "None" ? lv.type : "Active";
        return [w.name, w.designation, w.rc, String(activeCount), leaveStatus];
    });

    doc.autoTable({
        head, body,
        startY: startY + 3,
        theme: "grid",
        styles: { fontSize: 8, cellPadding: 2 },
        headStyles: { fillColor: [31, 41, 55], textColor: 255, fontStyle: "bold" },
        alternateRowStyles: { fillColor: [249, 250, 251] },
        columnStyles: {
            0: { cellWidth: 50, fontStyle: "bold" },
            1: { cellWidth: 54 },
            2: { cellWidth: 18, halign: "center" },
            3: { cellWidth: 24, halign: "center" },
            4: { cellWidth: 36, halign: "center" }
        },
        didParseCell: (data) => {
            if (data.section === "body" && data.column.index === 4) {
                const val = data.cell.raw;
                if (val === "Active") {
                    data.cell.styles.fillColor = [220, 252, 231];
                    data.cell.styles.textColor = [22, 101, 52];
                } else {
                    data.cell.styles.fillColor = [254, 226, 226];
                    data.cell.styles.textColor = [127, 29, 29];
                    data.cell.styles.fontStyle = "bold";
                }
            }
        },
        margin: { left: 14, right: 14 }
    });

    return doc.lastAutoTable.finalY + 8;
}

// ─── Quick Summary (1-pager) ────────────────────────────────────
window.exportQuickPDF = function() {
    const jsPDF = _pdfCheckJsPDF();
    if (!jsPDF) return;
    showToast("⚡ Building Quick Summary...");

    try {
        const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
        const stats = _pdfBuildStats();

        _pdfAddHeader(doc, "QUICK SUMMARY");
        let y = _pdfAddMeta(doc, 28);
        y = _pdfKPIBoxes(doc, y, stats);

        // Compact task breakdown
        doc.setFont("helvetica", "bold");
        doc.setFontSize(10);
        doc.setTextColor(31, 41, 55);
        doc.text("TASK STATUS BREAKDOWN", 14, y);
        y += 2;

        doc.autoTable({
            head: [["Total", "Ongoing", "Onhold", "Pending", "Completed", "Cancelled"]],
            body: [[stats.tOn + stats.tOh + stats.tP + stats.tC + stats.tCn, stats.tOn, stats.tOh, stats.tP, stats.tC, stats.tCn]],
            startY: y + 2,
            theme: "grid",
            styles: { fontSize: 10, cellPadding: 3, halign: "center", fontStyle: "bold" },
            headStyles: { fillColor: [31, 41, 55], textColor: 255 },
            columnStyles: {
                0: { textColor: [31, 41, 55] },
                1: { textColor: [10, 132, 255] },
                2: { textColor: [255, 149, 0] },
                3: { textColor: [107, 114, 128] },
                4: { textColor: [50, 215, 75] },
                5: { textColor: [255, 59, 48] }
            },
            margin: { left: 14, right: 14 }
        });
        y = doc.lastAutoTable.finalY + 10;

        // Urgent/Critical only (the important ones on a 1-pager)
        const urgentCritical = orders.filter(o => (o.priority === "Urgent" || o.priority === "Critical") && o.overallProgress < 100)
            .sort((a,b) => a.priority === "Critical" ? -1 : 1);
        y = _pdfWOTable(doc, y, "🚨 URGENT & CRITICAL ACTIVE WOs", urgentCritical, [255, 59, 48]);

        _pdfAddFooter(doc);
        doc.save(`EES_QuickSummary_${new Date().toISOString().slice(0,10)}.pdf`);
        showToast("✅ Quick Summary downloaded!");
    } catch (err) {
        console.error(err);
        showToast("❌ PDF error: " + err.message, true);
    }
};

// ─── Full Report (multi-page) ──────────────────────────────────
window.exportFullPDF = function() {
    const jsPDF = _pdfCheckJsPDF();
    if (!jsPDF) return;
    showToast("📄 Building Full Report...");

    try {
        const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
        const stats = _pdfBuildStats();

        _pdfAddHeader(doc, "FULL REPORT");
        let y = _pdfAddMeta(doc, 28);
        y = _pdfKPIBoxes(doc, y, stats);

        // Task status breakdown
        doc.setFont("helvetica", "bold");
        doc.setFontSize(10);
        doc.setTextColor(31, 41, 55);
        doc.text("TASK STATUS BREAKDOWN", 14, y);
        y += 2;

        doc.autoTable({
            head: [["Total", "Ongoing", "Onhold", "Pending", "Completed", "Cancelled"]],
            body: [[stats.tOn + stats.tOh + stats.tP + stats.tC + stats.tCn, stats.tOn, stats.tOh, stats.tP, stats.tC, stats.tCn]],
            startY: y + 2,
            theme: "grid",
            styles: { fontSize: 10, cellPadding: 3, halign: "center", fontStyle: "bold" },
            headStyles: { fillColor: [31, 41, 55], textColor: 255 },
            columnStyles: {
                0: { textColor: [31, 41, 55] },
                1: { textColor: [10, 132, 255] },
                2: { textColor: [255, 149, 0] },
                3: { textColor: [107, 114, 128] },
                4: { textColor: [50, 215, 75] },
                5: { textColor: [255, 59, 48] }
            },
            margin: { left: 14, right: 14 }
        });
        y = doc.lastAutoTable.finalY + 10;

        // Urgent/Critical
        const urgentCritical = orders.filter(o => (o.priority === "Urgent" || o.priority === "Critical") && o.overallProgress < 100)
            .sort((a,b) => a.priority === "Critical" ? -1 : 1);
        if (y > 230) { doc.addPage(); _pdfAddHeader(doc, "FULL REPORT"); y = 28; }
        y = _pdfWOTable(doc, y, "🚨 URGENT & CRITICAL", urgentCritical, [255, 59, 48]);

        // Ongoing
        const ongoing = orders.filter(o => o.tasks.some(t => t.status === "Ongoing") && o.overallProgress < 100);
        if (y > 220) { doc.addPage(); _pdfAddHeader(doc, "FULL REPORT"); y = 28; }
        y = _pdfWOTable(doc, y, "🟢 ONGOING WORK ORDERS", ongoing, [50, 215, 75]);

        // Completed
        const completed = orders.filter(o => o.overallProgress === 100);
        if (y > 220) { doc.addPage(); _pdfAddHeader(doc, "FULL REPORT"); y = 28; }
        y = _pdfWOTable(doc, y, "✅ COMPLETED WORK ORDERS", completed, [22, 163, 74]);

        // Team workload
        if (y > 200) { doc.addPage(); _pdfAddHeader(doc, "FULL REPORT"); y = 28; }
        y = _pdfTeamSection(doc, y);

        _pdfAddFooter(doc);
        doc.save(`EES_FullReport_${new Date().toISOString().slice(0,10)}.pdf`);
        showToast("✅ Full Report downloaded!");
    } catch (err) {
        console.error(err);
        showToast("❌ PDF error: " + err.message, true);
    }
};

// Backward compatibility — old references still work
window.exportDashboardPDF = window.exportFullPDF;

// ─── JSON Text Export ────────────────────────────────────────────
window.exportJsonText = function() {
    if (!orders.length) {
        showToast("No WO data to export.");
        return;
    }

    try {
        const payload = {
            exportedAt: new Date().toISOString(),
            exportedBy: {
                name: currentUser?.name || "System",
                rc: currentUser?.rc || "",
                role: currentUser?.role || ""
            },
            lastImport: lastUpdate || "Never",
            summary: {
                totalWOs: orders.length,
                activeWOs: orders.filter(o => Number(o.overallProgress || 0) < 100).length,
                completedWOs: orders.filter(o => Number(o.overallProgress || 0) === 100).length,
                urgentCriticalWOs: orders.filter(o => ["Urgent", "Critical"].includes(normalizePriority(o.priority)) && Number(o.overallProgress || 0) < 100).length,
                ongoingWOs: orders.filter(o => Number(o.overallProgress || 0) < 100 && (o.tasks || []).some(t => t.status === "Ongoing")).length
            },
            workOrders: orders,
            workers: WORKERS,
            leaves: workerLeaves
        };

        const jsonText = JSON.stringify(payload, null, 2);
        const blob = new Blob([jsonText], { type: "text/plain;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `EES_WO_Database_${new Date().toISOString().slice(0,10)}.txt`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        showToast("✅ JSON Text downloaded!");
    } catch (err) {
        console.error(err);
        showToast("❌ JSON Text export failed: " + err.message, true);
    }
};

// ── Excel Import ──────────────────────────────────────────────────

window.processExcel = async function(file) {
    if (!requireAdmin("import Excel data")) return;
    if(!window.XLSX){showToast("⏳ Excel library loading...", true);return;}
    if(!isSupportedExcelFile(file)){showToast("⚠️ Please select a valid .xlsx or .xls file.", true);return;}
    if(file.size > 25 * 1024 * 1024){showToast("⚠️ Excel file is too large. Please keep it under 25 MB.", true);return;}

    const reader = new FileReader();
    reader.onload = async (e) => {
        showLoad("Reading Excel file...");
        try {
            if (!e.target || !e.target.result) throw new Error("Excel file could not be read.");

            const wb = XLSX.read(new Uint8Array(e.target.result), { type:'array', cellDates:true });
            const ws = getFirstWorksheet(wb);
            const rows = getValidatedExcelRows(ws);
            const columns = getExcelImportColumns(rows);

            const eew = rows.filter(row => {
                const unit = String(row["UNIT"] || "").trim().toUpperCase();
                const woNumber = String(row["WO NUMBER"] || row["WO ID"] || "").trim().toUpperCase();
                return unit === "EEW" || woNumber.includes("-EEW");
            });

            if(!eew.length){hideLoad();showToast("⚠️ No EEW records found.");return;}

            let added = 0, updated = 0, skipped = 0;
            const changed = new Map();
            const ordersById = buildOrderIndex(orders);
            const taskIndexesByWorkOrderId = new Map();

            eew.forEach(row => {
                const wid = String(row["WO NUMBER"] || row["WO ID"] || "").trim();
                if(!wid || wid === "0") { skipped++; return; }
                if(!hasUsableTaskData(row, columns)) { skipped++; return; }

                const ed = formatForInput(row["TASK DATE"]) || formatForInput(row[columns.startDate]) || "";
                let wo = ordersById.get(wid);

                if(!wo) {
                    wo = {
                        id: wid,
                        sr: String(row["SR NUMBER"] || "").trim(),
                        svo: String(row["SVO NUMBER"] || "").trim(),
                        asset: String(row["ASSET/SERVICE NAME"] || "TBA").trim() || "TBA",
                        date: ed,
                        priority: "Minor",
                        overallProgress: 0,
                        assignees: [],
                        tasks: []
                    };
                    orders.push(wo);
                    ordersById.set(wid, wo);
                } else if(isBadDate(wo.date) && ed) {
                    wo.date = ed;
                }

                const taskNo = String(row["TASK NO:"] || `Task-${wo.tasks.length + 1}`).trim();
                const incomingTask = makeImportedTask(row, columns, taskNo);
                const taskIndex = getCachedTaskIndex(wo, taskIndexesByWorkOrderId);
                const existingTask = taskIndex.get(taskNo);

                if(!existingTask) {
                    wo.tasks.push(incomingTask);
                    taskIndex.set(taskNo, incomingTask);
                    added++;
                    changed.set(wo.id, wo);
                    return;
                }

                if(mergeImportedTask(existingTask, incomingTask, row, columns)) {
                    updated++;
                    changed.set(wo.id, wo);
                }
            });

            changed.forEach(wo => {
                const normalized = normalizeWorkOrder(wo);
                Object.assign(wo, normalized);
            });

            if(changed.size > 0){
                showLoad(`Saving ${changed.size} work orders to Firebase...`);
                await dbBatchSave([...changed.values()]);
                lastUpdate = new Date().toLocaleString("en-US", { dateStyle:"medium", timeStyle:"short" });
                await dbSaveLastUpdate(lastUpdate);
                await auditLog("EXCEL_IMPORTED", { changedWorkOrders: changed.size, addedTasks: added, updatedTasks: updated, skippedRows: skipped });
                const skippedMsg = skipped ? `, skipped ${skipped} incomplete rows` : "";
                showToast(`✅ Added ${added} new tasks, updated ${updated} existing${skippedMsg}`);
                setView('orders');
            } else {
                const skippedMsg = skipped ? ` Skipped ${skipped} incomplete rows.` : "";
                showToast(`⚠️ No changes — database already up to date.${skippedMsg}`);
            }
        } catch(err) {
            console.error(err);
            showToast("❌ Error: " + err.message, true);
        }
        hideLoad(); renderApp();
    };
    reader.onerror = () => {
        hideLoad();
        showToast("❌ Could not read the selected Excel file.", true);
    };
    reader.readAsArrayBuffer(file);
};

window.handleFileSelect = e => { if(e.target.files[0]) processExcel(e.target.files[0]); };
window.handleDrop = e => { e.preventDefault(); e.currentTarget.classList.remove('drag'); if(e.dataTransfer.files[0]) processExcel(e.dataTransfer.files[0]); };

// ── Render helpers ────────────────────────────────────────────────

function PTagHTML(p) {
    const priority = normalizePriority(p);
    const c = PRI_CFG[priority] || PRI_CFG["Minor"];
    return `<span class="ptag" style="color:${c.color};border:1px solid ${c.color}60;">${html(priority)}</span>`;
}

function getPriorityColor(priority) {
    return (PRI_CFG[normalizePriority(priority)] || PRI_CFG["Minor"]).color;
}

function renderWOCardHTML(order, options = {}) {
    const isCompleted = Boolean(options.isCompleted);
    const isOngoing = Boolean(options.isOngoing);
    const clickAction = options.clickAction || `selectOrder(${jsArg(order.id)})`;
    const extraClass = options.extraClass || "";
    const progressPrefix = options.progressPrefix || "";
    const borderStyle = options.borderStyle || `border-left-color:${getPriorityColor(order.priority)}`;
    const progressStyle = isCompleted ? ' style="color:var(--green)"' : "";

    return `<div class="wo-card ${extraClass}" style="${attr(borderStyle)}" onclick="${clickAction}">
        <div class="wo-col left"><div class="wo-id">${html(order.id)}</div><div class="wo-asset">${html(order.asset)}</div></div>
        <div class="wo-col center"><div class="wo-date">${html(formatDateNice(order.date))}</div><div class="wo-tasks">${clampNumber(order.tasks?.length || 0, 0, 999999)} Tasks</div></div>
        <div class="wo-col right">${PTagHTML(order.priority)}<div class="wo-prog"${progressStyle}>${html(progressPrefix)}${clampNumber(order.overallProgress || 0, 0, 100)}%</div></div>
    </div>`;
}

function getSidebarHTML() {
    const iv = isViewerUser();
    let onlineHtml = "";
    if (isAdminUser()) {
        onlineHtml = `<div class="sidebar-presence"><span class="sidebar-presence-name">${html(viewerPresence.name)}</span><span id="viewer-status">${getPresenceStatusHTML()}</span></div>`;
    }
    return `
    <aside class="sidebar ${isSidebarOpen ? 'show' : ''}">
        <div class="logo">
            <div class="logo-title">EES WO CONTROL<br>DATABASE</div>
        </div>
        <div class="sidebar-main">
            <div class="nav-grp"><div class="nav-grp-lbl">Main</div>
                <div class="nav-item ${view==='dashboard'?'active':''}" onclick="setView('dashboard')"><span>📊</span> Dashboard</div>
                <div class="reports-menu">
                    <button class="nav-item reports-menu-btn ${['orders','urgent','ongoing','completed'].includes(view)?'active':''}" type="button" onclick="toggleWOViewsMenu(event)">
                        <span>📋</span><span class="reports-menu-text">Work Orders</span><span class="reports-caret">▾</span>
                    </button>
                    <div class="reports-menu-panel ${isWOViewsMenuOpen ? 'show' : ''}">
                        <button type="button" class="${view==='orders'?'active':''}" onclick="setView('orders')"><span>📋</span> Active Orders</button>
                        <button type="button" class="${view==='urgent'?'active':''}" onclick="setView('urgent')"><span>🚨</span> Urgent & Critical</button>
                        <button type="button" class="${view==='ongoing'?'active':''}" onclick="setView('ongoing')"><span>🟢</span> Ongoing WOs</button>
                        <button type="button" class="${view==='completed'?'active':''}" onclick="setView('completed')"><span>✅</span> WO Completed</button>
                    </div>
                </div>
            </div>
            <div class="nav-grp"><div class="nav-grp-lbl">Team</div>
                <div class="nav-item ${view==='workers'?'active':''}" onclick="setView('workers')"><span>👷</span> EES Team</div>
            </div>
            <div class="nav-grp"><div class="nav-grp-lbl">Reports & Data</div>
                ${!iv?`<div class="nav-item ${view==='upload'?'active':''}" onclick="setView('upload')"><span>⬆️</span> Manage Data</div>`:''}
                <div class="reports-menu">
                    <button class="nav-item reports-menu-btn" type="button" onclick="toggleReportsMenu(event)">
                        <span>📦</span><span class="reports-menu-text">Exports</span><span class="reports-caret">▾</span>
                    </button>
                    <div class="reports-menu-panel ${isReportsMenuOpen ? 'show' : ''}">
                        <button type="button" onclick="runReportAction('excel')"><span>⬇️</span> Export Excel</button>
                        <button type="button" onclick="runReportAction('quickPdf')"><span>⚡</span> Quick PDF</button>
                        <button type="button" onclick="runReportAction('fullPdf')"><span>📄</span> Full PDF</button>
                        <button type="button" onclick="runReportAction('jsonText')"><span>{ }</span> JSON Text</button>
                    </div>
                </div>
            </div>
        </div>
        <div class="sidebar-foot">
            ${onlineHtml}
            <div class="user-card">
                <div class="avc">${html(ini(currentUser.name))}</div>
                <div><div class="user-name">${html(currentUser.name)}</div><div class="user-role">${html(currentUser.designation)}</div></div>
            </div>
            <div class="theme-row">
                <span style="font-size:12px;color:var(--muted);font-weight:600;">Light Mode</span>
                <label class="switch"><input type="checkbox" onchange="toggleTheme()" ${isLightMode?'checked':''}><span class="slider"></span></label>
            </div>
            <div class="logout-text" onclick="handleLogout()">Logout</div>
        </div>
    </aside>`;
}

// ── Main render ───────────────────────────────────────────────────

function renderApp() {
    const root = document.getElementById("app-root");

    if(!loggedIn) {
        root.innerHTML=`<div class="login-wrap"><div class="login-card">
            <div class="login-title">EES WO CONTROL DATABASE</div>
            <div class="login-subtitle">MTCC</div>
            <div class="login-inp-group">
                <label class="login-lbl">Email</label>
                <input id="login-u" class="login-inp" type="text" placeholder="your.name@mtcc.com.mv" onkeydown="if(event.key==='Enter')handleLogin()">
                <label class="login-lbl">Password</label>
                <div class="pwd-container">
                    <input id="login-p" class="login-inp" type="password" placeholder="Enter password" style="padding-right:40px" onkeydown="if(event.key==='Enter')handleLogin()">
                    <div class="pwd-toggle" onclick="togglePassword()">
                        <svg id="eye-icon" viewBox="0 0 24 24" width="18" height="18" stroke="var(--muted)" stroke-width="2.5" fill="none" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"></path>
                            <line x1="1" y1="1" x2="23" y2="23"></line>
                        </svg>
                    </div>
                </div>
                ${loginErr?`<div class="login-err">⚠ ${html(loginErr)}</div>`:''}
            </div>
            <button class="login-btn" onclick="handleLogin()">Sign In →</button>
        </div></div>`;
        return;
    }

    const isViewer = isViewerUser();
    const dis = isViewer?'disabled':'';
    let contentHtml="", detailHtml="", topbarExtra="";

    // ── Dashboard ──
    if(view==="dashboard") {
        let onLeave=0;
        WORKERS.forEach(w=>{if(workerLeaves[w.rc]?.type&&workerLeaves[w.rc].type!=="None")onLeave++;});

        const activeCount = orders.filter(o=>o.overallProgress<100).length;
        const completedCount = orders.filter(o=>o.overallProgress===100).length;
        const urgentCriticalCount = orders.filter(o=>(o.priority==="Urgent"||o.priority==="Critical")&&o.overallProgress<100).length;
        const ongoingWOCount = orders.filter(o=>o.tasks.some(t=>t.status==="Ongoing")&&o.overallProgress<100).length;

        let tT=0,tOn=0,tOh=0,tP=0,tC=0,tCn=0;
        orders.forEach(o=>o.tasks.forEach(t=>{tT++;if(t.status==="Ongoing")tOn++;if(t.status==="Onhold")tOh++;if(t.status==="Pending")tP++;if(t.status==="Completed")tC++;if(t.status==="Cancelled")tCn++;}));

        contentHtml=`
        <div id="dashboard-content" style="display:flex;flex-direction:column;gap:30px;padding-bottom:20px;">
            <div><div class="sec-title">EES MANPOWER</div><div class="task-stats">
                <div class="ts-item"><div class="ts-val">${WORKERS.length}</div><div class="ts-lbl">Headcount</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--blue)">${WORKERS.length-onLeave}</div><div class="ts-lbl">On Duty</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--red)">${onLeave}</div><div class="ts-lbl">On Leave</div></div>
            </div></div>
            <div><div class="sec-title">WO Data</div><div class="task-stats">
                <div class="ts-item"><div class="ts-val">${orders.length}</div><div class="ts-lbl">Total WOs</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--blue)">${activeCount}</div><div class="ts-lbl">Active</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--green)">${completedCount}</div><div class="ts-lbl">Completed</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--red)">${urgentCriticalCount}</div><div class="ts-lbl">Urgent/Critical</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--orange)">${ongoingWOCount}</div><div class="ts-lbl">Ongoing WOs</div></div>
            </div></div>
            <div><div class="sec-title">Tasks Status</div><div class="task-stats">
                <div class="ts-item"><div class="ts-val">${tT}</div><div class="ts-lbl">Total</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--blue)">${tOn}</div><div class="ts-lbl">Ongoing</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--orange)">${tOh}</div><div class="ts-lbl">Onhold</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--muted)">${tP}</div><div class="ts-lbl">Pending</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--green)">${tC}</div><div class="ts-lbl">Completed</div></div>
                <div class="ts-item"><div class="ts-val" style="color:var(--red)">${tCn}</div><div class="ts-lbl">Cancelled</div></div>
            </div></div>
        </div>`;
    }

    // ── Work Order List Views ──
    else if(isWorkOrderView(view)) {
        const viewConfig = getWorkOrderViewConfig(view);
        const isc = viewConfig.isCompleted;
        const base = viewConfig.list;
        const state = getOrderViewState(view);
        const isActiveOrdersView = view === "orders";
        const assets = Array.from(new Set(base.map(o => o.asset))).filter(Boolean).sort();
        const filteredByAsset = isActiveOrdersView
            ? base.filter(o => state.filterAsset === "All" || o.asset === state.filterAsset)
            : base;
        const filtered = isActiveOrdersView
            ? filteredByAsset.filter(o => workOrderMatchesSearch(o, state.search))
            : filteredByAsset;
        const controlsHtml = isActiveOrdersView ? `
        <div style="display:grid;grid-template-columns:minmax(200px,320px) minmax(220px,1fr) auto;gap:12px;align-items:end;margin-bottom:16px;">
            <div><label style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:8px;display:block;font-weight:600;">Filter by Asset</label>
            <select class="filter-dropdown" style="margin-bottom:0;" onchange="setFilter(this.value)"><option value="All" ${state.filterAsset==="All"?"selected":""}>All Assets</option>${assets.map(a=>`<option value="${attr(a)}" ${state.filterAsset===a?"selected":""}>${html(a)}</option>`).join('')}</select></div>
            <div><label style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:8px;display:block;font-weight:600;">Search</label>
            <input id="active-order-search" class="input" type="search" value="${attr(state.search)}" placeholder="Search WO, asset, SR, SVO, assignee, task..." onkeydown="if(event.key==='Enter')applyOrderSearch()"></div>
            <div><button class="btn btn-primary" style="height:48px;" onclick="applyOrderSearch()">Search</button></div>
        </div>` : "";
        const cardsHtml = filtered.map(o => {
            const on = o.tasks.some(t => t.status === "Ongoing");
            return renderWOCardHTML(o, {
                isCompleted: isc,
                extraClass: on && !isc ? "ongoing-glow" : "",
                borderStyle: (!on || isc) ? `border-left-color:${getPriorityColor(o.priority)}` : ""
            });
        }).join('');
        contentHtml = `
        ${controlsHtml}
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;color:var(--muted);font-size:12px;font-weight:600;">
            <span>${filtered.length} ${isActiveOrdersView ? "matching " : ""}${viewConfig.matchLabel} work orders</span>
            ${isActiveOrdersView && state.search ? `<button class="btn btn-ghost" onclick="setOrderSearch('')">Clear Search</button>` : ''}
        </div>
        <div class="wo-list">${filtered.length===0?`<div style="text-align:center;color:var(--muted);padding:40px;">No ${viewConfig.emptyLabel} work orders${isActiveOrdersView ? " match your filters" : ""}.</div>`:''}
        ${cardsHtml}</div>`;
    }

    // ── Workers ──
    else if(view==="workers") {
        contentHtml=`<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:20px;">`+
        WORKERS.map(w=>{
            const wos=orders.filter(o=>o.assignees?.includes(w.rc));
            const act=wos.filter(o=>o.tasks.some(t=>t.status==="Ongoing")).length;
            const lv=workerLeaves[w.rc];
            const lvB=lv&&lv.type!=='None'?`<div style="background:var(--red);color:#fff;font-size:10px;font-weight:700;padding:4px 8px;border-radius:4px;margin-top:8px;display:inline-block;">🔴 ${html(lv.type)}${lv.from&&lv.to?` (${html(formatDateNice(lv.from))} – ${html(formatDateNice(lv.to))})`:''}</div>`:'';
            const md=isViewer?'':`<div class="menu-dots" onclick="event.stopPropagation();openLeaveModal(${jsArg(w.rc)})"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="18" x2="21" y2="18"/></svg></div>`;
            return`<div class="worker-card" onclick="selectWorker(${jsArg(w.rc)})">${md}<div style="display:flex;gap:12px;align-items:center;"><div class="avc" style="width:44px;height:44px;font-size:14px;">${html(ini(w.name))}</div><div><div style="font-weight:700;font-size:15px;color:var(--text);padding-right:28px;">${html(w.name)}</div><div style="font-size:12px;color:var(--muted);">${html(w.designation)} · RC ${html(w.rc)}</div></div></div>${lvB}<div style="margin-top:14px;padding-top:12px;border-top:1px solid var(--border);display:flex;justify-content:space-between;"><div style="font-size:15px;font-weight:700;">${wos.length} <span style="font-size:10px;color:var(--muted);">Total WOs</span></div><div style="font-size:11px;font-weight:600;color:var(--blue);">${act} Ongoing</div></div></div>`;
        }).join('')+`</div>`;
    }

    // ── Upload ──
    else if(view==="upload") {
        contentHtml=isViewer?`<div style="text-align:center;padding:50px;color:var(--muted);">No permission to view this page.</div>`:`
        <div style="max-width:520px;margin:0 auto;">
            <div class="sec-title" style="text-align:center;">Import / Manage Data</div>
            <div style="text-align:center;color:var(--muted);font-size:13px;line-height:1.6;margin:12px 0 16px;">Select the <strong>Scope Master</strong> Excel file.<br>Only <strong>EEW</strong> records will be imported and saved to Firebase.</div>
            <div style="text-align:center;font-size:13px;margin-bottom:20px;">Last import: <span style="color:var(--blue);font-weight:600;">${html(lastUpdate)}</span></div>
            <div class="upload-zone" ondragover="event.preventDefault();this.classList.add('drag')" ondragleave="this.classList.remove('drag')" ondrop="handleDrop(event)" onclick="document.getElementById('f-in').click()">
                <input id="f-in" type="file" accept=".xlsx,.xls" style="display:none" onchange="handleFileSelect(event)">
                <div style="font-size:46px;margin-bottom:12px;">📂</div>
                <div style="font-size:16px;font-weight:600;margin-bottom:4px;">Click to Browse</div>
                <div style="font-size:12px;color:var(--muted);">Or drag and drop your Excel file here</div>
            </div>
            <div style="margin-top:28px;padding-top:22px;border-top:1px dashed var(--border);">
                <button class="btn" style="background:var(--red);color:#fff;width:100%;padding:14px;font-size:14px;justify-content:center;" onclick="clearDatabase()">🗑️ Clear Entire Database</button>
            </div>
        </div>`;
    }

    // ── Leave Modal ──
    if(leaveModalWorker&&view==="workers") {
        const w=WORKERS.find(x=>x.rc===leaveModalWorker);
        if(!w) { leaveModalWorker=null; }
        else {
            const lv=workerLeaves[w.rc]||{type:'None',from:'',to:''};
            const leaveType = String(lv.type || 'None');
            const sd=(leaveType&&leaveType!=='None')?'flex':'none';
            const ft=isViewer?`<button class="btn btn-ghost" onclick="closeLeaveModal()">Close</button>`:`<button class="btn btn-ghost" onclick="closeLeaveModal()">Cancel</button><button class="btn btn-primary" onclick="saveLeaveUpdates(${jsArg(w.rc)})">💾 Save</button>`;
            detailHtml=`<div class="overlay" onclick="if(event.target===this)closeLeaveModal()"><div class="modal" style="max-width:400px;"><div class="modal-head"><div style="font-family:'Rajdhani',sans-serif;font-size:20px;font-weight:700;">Leave Status</div><button onclick="closeLeaveModal()" style="background:none;border:none;font-size:22px;color:var(--muted);cursor:pointer;">✕</button></div><div class="modal-body"><div style="font-weight:700;font-size:16px;margin-bottom:16px;">${html(w.name)}</div><label class="plbl">Leave Type</label><select id="l-type" class="select" ${dis} style="margin-bottom:14px;" onchange="document.getElementById('l-dates').style.display=this.value==='None'?'none':'flex'"><option value="None" ${leaveType==='None'?'selected':''}>Active (No Leave)</option><option value="Annual Leave" ${leaveType==='Annual Leave'?'selected':''}>Annual Leave</option><option value="Sick Leave" ${leaveType==='Sick Leave'?'selected':''}>Sick Leave</option><option value="FRL" ${leaveType==='FRL'?'selected':''}>FRL</option></select><div id="l-dates" style="display:${sd};flex-direction:column;gap:14px;"><div><label class="plbl">From</label><input type="date" id="l-from" class="input" ${dis} value="${attr(lv.from||'')}"></div><div><label class="plbl">To</label><input type="date" id="l-to" class="input" ${dis} value="${attr(lv.to||'')}"></div></div></div><div class="modal-foot">${ft}</div></div></div>`;
        }
    }

    // ── Worker Profile Modal ──
    if(selectedWorker&&!leaveModalWorker&&view==="workers") {
        const w=WORKERS.find(x=>x.rc===selectedWorker);
        if(!w) { selectedWorker=null; }
        else {
            const wos=orders.filter(o=>o.assignees?.includes(w.rc));
            const act=wos.filter(o=>o.tasks.some(t=>t.status==="Ongoing")).length;
            const wh=wos.length===0?`<div style="font-size:13px;color:var(--muted);text-align:center;padding:18px;background:var(--bg);border-radius:8px;">No assigned work orders.</div>`:`<div style="display:grid;gap:8px;">`+wos.map(o=>`<div style="background:var(--bg);border:1px solid var(--border);padding:12px;border-radius:8px;display:flex;justify-content:space-between;align-items:center;cursor:pointer;" onclick="openWOFromWorker(${jsArg(o.id)})"><div><div style="font-weight:700;">${html(o.id)}</div><div style="font-size:11px;color:var(--muted);">${html(o.asset)}</div></div><div style="text-align:right;">${PTagHTML(o.priority)}<div style="font-size:11px;font-weight:700;margin-top:4px;">${clampNumber(o.overallProgress||0,0,100)}%</div></div></div>`).join('')+`</div>`;
            detailHtml=`<div class="overlay" onclick="if(event.target===this)closeWorkerModal()"><div class="modal" style="max-width:500px;"><div class="modal-head"><div style="font-family:'Rajdhani',sans-serif;font-size:20px;font-weight:700;">Employee Profile</div><button onclick="closeWorkerModal()" style="background:none;border:none;font-size:22px;color:var(--muted);cursor:pointer;">✕</button></div><div class="modal-body"><div style="display:flex;align-items:center;gap:15px;margin-bottom:22px;"><div class="avc" style="width:56px;height:56px;font-size:18px;">${html(ini(w.name))}</div><div><div style="font-weight:700;font-size:20px;">${html(w.name)}</div><div style="font-size:13px;color:var(--blue);font-weight:600;">${html(w.designation)}</div></div></div><div class="d-meta" style="grid-template-columns:1fr;"><div style="display:flex;justify-content:space-between;"><span>RC Number</span><span>${html(w.rc)}</span></div><div style="display:flex;justify-content:space-between;"><span>Phone</span><span>${html(w.phone||'—')}</span></div><div style="display:flex;justify-content:space-between;"><span>Email</span><span>${html(w.email||'—')}</span></div></div><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;"><label class="plbl" style="margin:0;">Assigned WOs (${wos.length})</label><span style="font-size:11px;font-weight:600;color:var(--blue);">${act} Ongoing</span></div>${wh}</div></div></div>`;
        }
    }

    // ── WO Detail Modal ──
    if(selectedWO&&isWorkOrderView(view)) {
        const wo=orders.find(o=>o.id===selectedWO);
        if(wo) {
            const tags=wo.assignees.map(rc=>{
                const w=WORKERS.find(x=>x.rc===rc);
                const displayName = w ? String(w.name).split(" ")[0] : rc;
                const rb=isViewer?'':`<span class="atag-x" onclick="removeAssignee(${jsArg(rc)})">✕</span>`;
                return`<div class="atag">${html(displayName)} ${rb}</div>`;
            }).join('');
            const ft=isViewer?`<button class="btn btn-ghost" onclick="selectOrder(null)">Close</button>`:`<button class="btn btn-ghost" onclick="selectOrder(null)">Cancel</button><button class="btn btn-primary" style="padding:10px 30px;" onclick="saveWOUpdates(${jsArg(wo.id)})">💾 Save</button>`;
            const assigneeOptions = WORKERS
                .filter(w=>!wo.assignees.includes(w.rc))
                .map(w=>`<option value="${attr(w.rc)}">${html(w.name)} (${html(w.designation)})</option>`)
                .join('');
            detailHtml=`<div class="overlay" onclick="if(event.target===this)selectOrder(null)"><div class="modal"><div class="modal-head"><div style="font-family:'Rajdhani',sans-serif;font-size:24px;font-weight:700;">${html(wo.id)}</div><button onclick="selectOrder(null)" style="background:none;border:none;font-size:22px;color:var(--muted);cursor:pointer;">✕</button></div><div class="modal-body">
            <div class="d-meta"><div>Asset: <span>${html(wo.asset)}</span></div><div>Date: <span>${html(formatDateNice(wo.date))}</span></div><div>SR No: <span>${html(wo.sr||'—')}</span></div><div>SVO No: <span>${html(wo.svo||'—')}</span></div></div>
            <div style="margin-bottom:22px;"><label class="plbl">WO Priority</label><select id="wo-priority" class="select" ${dis}><option value="Minor" ${wo.priority==="Minor"?'selected':''}>Minor</option><option value="Major" ${wo.priority==="Major"?'selected':''}>Major</option><option value="Urgent" ${wo.priority==="Urgent"?'selected':''}>Urgent</option><option value="Critical" ${wo.priority==="Critical"?'selected':''}>Critical</option></select></div>
            <div class="assign-wrap"><label class="plbl">Assigned Team</label>${!isViewer?`<select class="select" onchange="addAssignee(this.value);this.value='';" style="margin-bottom:10px;"><option value="">+ Assign employee...</option>${assigneeOptions}</select>`:''}<div class="assign-tags">${tags}</div></div>
            <div><label class="plbl">Tasks (${wo.tasks.length})</label>${wo.tasks.length===0?'<div style="font-size:12px;color:var(--muted);">No tasks for this WO.</div>':''}
            ${wo.tasks.map((t,i)=>`<div class="task-card"><div class="task-head">${html(t.taskNo)}: ${html(t.details)}</div><div class="task-body"><div><label class="plbl">Status</label><select id="t-stat-${i}" class="select" ${dis}><option ${t.status==="Pending"?'selected':''}>Pending</option><option ${t.status==="Ongoing"?'selected':''}>Ongoing</option><option ${t.status==="Onhold"?'selected':''}>Onhold</option><option ${t.status==="Completed"?'selected':''}>Completed</option><option ${t.status==="Cancelled"?'selected':''}>Cancelled</option></select></div><div><label class="plbl">Progress %</label><input type="number" id="t-prog-${i}" class="input" ${dis} value="${attr(clampNumber(t.progress,0,100))}" min="0" max="100"></div><div><label class="plbl">Start Date</label><input type="date" id="t-start-${i}" class="input" ${dis} value="${attr(t.startDate)}">${t.startDate?`<div style="font-size:10px;color:var(--muted);margin-top:3px;">${html(formatDateNice(t.startDate))}</div>`:''}</div><div><label class="plbl">Complete Date</label><input type="date" id="t-end-${i}" class="input" ${dis} value="${attr(t.completeDate)}">${t.completeDate?`<div style="font-size:10px;color:var(--muted);margin-top:3px;">${html(formatDateNice(t.completeDate))}</div>`:''}</div><div class="task-full"><label class="plbl">Remarks</label><input type="text" id="t-rem-${i}" class="input" ${dis} value="${attr(t.remarks)}" placeholder="Notes..."></div></div></div>`).join('')}
            </div></div><div class="modal-foot">${ft}</div></div></div>`;
        }
    }

    root.innerHTML=`<div class="app"><div class="sidebar-overlay ${isSidebarOpen?'show':''}" onclick="toggleSidebar()"></div>${getSidebarHTML()}<div class="main"><div class="topbar"><div style="display:flex;align-items:center;"><button class="menu-btn" onclick="toggleSidebar()">☰</button><div class="page-title">${pageLabels[view]||""}</div></div>${topbarExtra}</div><div class="body-wrap"><div class="content">${contentHtml}</div>${detailHtml}</div></div></div>`;
    publishAppView();
}

// ── Init: check if user already signed in ─────────────────────────
auth.onAuthStateChanged(async user => {
    if(user) {
        const profile = await loadUserProfile(user);
        if(profile) {
            currentUser=profile; loggedIn=true; view="dashboard"; publishCurrentUser(); publishAppView();
            showLoad("Loading your data...");
            try { await dbLoadAll(); await startUserSession(); } catch(e) { console.error(e); }
            hideLoad(); renderApp();
            startRealtimeSync();        } else {
            currentUser=null; loggedIn=false; view="dashboard"; publishCurrentUser(); publishAppView();
            hideLoad(); renderApp();
        }
    } else {
        currentUser=null; loggedIn=false; view="dashboard"; publishCurrentUser(); publishAppView();
        hideLoad(); renderApp();
    }
});
