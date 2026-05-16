// ============================================================
//  app.js — 113級3班 成績追蹤系統（主程式）
// ============================================================

// ── Firebase 初始化 ──────────────────────────────────────────
let db, auth;
try {
  firebase.initializeApp(FIREBASE_CONFIG);
  db   = firebase.firestore();
  auth = firebase.auth();
} catch(e) {
  console.warn("Firebase init failed:", e.message);
}


// ── 自動重新登入（重新整理後不需重新輸入密碼）────────────────
if (auth) {
  auth.onAuthStateChanged(async user => {
    const mainPage = document.getElementById("main-page");
    if (mainPage && !mainPage.classList.contains("hidden")) return;
    if (!user) {
      const lp = document.getElementById("login-page");
      if (lp) lp.classList.remove("hidden");
      return;
    }
    try {
      const linfo = document.getElementById("linfo");
      if (linfo) { linfo.textContent = "⏳ 自動登入中..."; linfo.style.display = "block"; }
      let teacherDoc = null;
      const byUid = await db.collection("teachers").doc(user.uid).get();
      if (byUid.exists) {
        teacherDoc = byUid.data();
      } else {
        const email = user.email || "";
        const acc = email.includes("@") ? email.split("@")[0] : email;
        const byAcc = await db.collection("teachers").where("account", "==", acc).limit(1).get();
        if (!byAcc.empty) teacherDoc = byAcc.docs[0].data();
      }
      if (!teacherDoc) { auth.signOut(); return; }
      PREFIX = teacherDoc.classPrefix + "_";
      CURRENT_CLASS.prefix      = teacherDoc.classPrefix;
      CURRENT_CLASS.className   = teacherDoc.className   || CLASS_NAME;
      CURRENT_CLASS.classYear   = teacherDoc.classYear   || CLASS_YEAR;
      CURRENT_CLASS.teacherName = teacherDoc.teacherName || "";
      CURRENT_CLASS.isAdmin     = teacherDoc.isAdmin === true;
      if (Array.isArray(teacherDoc.subjects) && teacherDoc.subjects.length > 0) {
        ACTIVE_SUBJECTS = teacherDoc.subjects;
      } else {
        ACTIVE_SUBJECTS = (typeof SUBJECTS !== "undefined") ? [...SUBJECTS] : [];
      }
      if (Array.isArray(teacherDoc.exams) && teacherDoc.exams.length > 0) {
        ACTIVE_EXAMS = teacherDoc.exams;
      } else {
        ACTIVE_EXAMS = (typeof EXAMS !== "undefined") ? [...EXAMS] : [];
      }
      if (CURRENT_CLASS.isAdmin) {
        try {
          const allSnap = await db.collection("teachers").get();
          CURRENT_CLASS.allClasses = allSnap.docs.map(d => ({ id: d.id, ...d.data() }));
        } catch(e) { CURRENT_CLASS.allClasses = []; }
      }
      S.cur     = teacherDoc.account || "";
      S.isAdmin = CURRENT_CLASS.isAdmin;
      await loadAllData();
      showMainPage();
    } catch(e) {
      console.warn("自動登入失敗：", e.message);
      const lp = document.getElementById("login-page");
      if (lp) lp.classList.remove("hidden");
    }
  });
}

// 動態科目與段考（登入後可被 Firestore 覆蓋）
// 預設值來自 config.js，各班可在 teachers 文件裡自訂
let ACTIVE_SUBJECTS = (typeof SUBJECTS !== "undefined") ? [...SUBJECTS] : [];
let ACTIVE_EXAMS    = (typeof EXAMS    !== "undefined") ? [...EXAMS]    : [];
function getClassName() { return CURRENT_CLASS.className || CLASS_NAME; }
function getClassYear() { return CURRENT_CLASS.classYear || CLASS_YEAR; }
let PREFIX = (typeof COLLECTION_PREFIX !== "undefined") ? COLLECTION_PREFIX + "_" : "grade_";

// 目前登入老師的班級資訊（多班級支援）
let CURRENT_CLASS = {
  prefix:    typeof COLLECTION_PREFIX !== "undefined" ? COLLECTION_PREFIX : "grade",
  className: typeof CLASS_NAME  !== "undefined" ? CLASS_NAME  : "班級",
  classYear: typeof CLASS_YEAR  !== "undefined" ? CLASS_YEAR  : "",
  teacherName: "",
  isAdmin:   false,   // 管理者可切換班級
  allClasses: []      // 管理者專用：所有班級清單
};

function col(name) {
  const prefixed = ["scores", "students", "config"];
  return db.collection(prefixed.includes(name) ? PREFIX + name : name);
}

// ── 狀態 ─────────────────────────────────────────────────────
const S = {
  cur: null, isAdmin: false,
  students: [],
  scores: {},
  activePage: "overview",
  analysisStudentId: null,
  reportStudentId: null,
  teacherComments: {},
  studentMemos: {},       // { studentId: "備忘錄內容" } — 學生個別私人備忘
  examSubjectCountCache: {},
};

// ── 工具 ──────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const show = id => { const e=$(id); if(e) e.classList.remove("hidden"); };
const hide = id => { const e=$(id); if(e) e.classList.add("hidden"); };

let toastTimer;
function showToast(msg) {
  const t = $("toast"); if(!t) return;
  t.textContent = msg; t.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove("show"), 2800);
}

function fmt(v, dec=1) {
  return (v !== null && v !== undefined && v !== "") ? parseFloat(v).toFixed(dec) : "—";
}
function avg(arr) {
  const v = arr.filter(x => x !== null && x !== undefined && x !== "");
  return v.length ? v.reduce((a,b) => a + parseFloat(b), 0) / v.length : null;
}
// ── 動態滿分計算（依實際填寫科目數）────────────────────────
function getExamSubjectCount(examId) {
  if (S.examSubjectCountCache[examId] !== undefined) return S.examSubjectCountCache[examId];
  const filled = ACTIVE_SUBJECTS.filter(sub =>
    S.students.some(st => {
      const v = getScores(st.id, examId)[sub];
      return v !== undefined && v !== "";
    })
  );
  const result = filled.length || ACTIVE_SUBJECTS.length;
  S.examSubjectCountCache[examId] = result;
  return result;
}
// 成績變更時清除相關快取
function clearExamSubjectCountCache(examId) {
  if (examId) delete S.examSubjectCountCache[examId];
  else S.examSubjectCountCache = {};
}
function getStudentMaxScore(studentId, examId) {
  const sc = getScores(studentId, examId);
  const filledCount = ACTIVE_SUBJECTS.filter(s => sc[s] !== undefined && sc[s] !== "").length;
  return filledCount * 100 || ACTIVE_SUBJECTS.length * 100;
}
function getExamMaxScore(examId) {
  return getExamSubjectCount(examId) * 100;
}

function getTotal(sc) {
  const vals = ACTIVE_SUBJECTS.map(s => sc && sc[s]).filter(v => v !== null && v !== undefined && v !== "");
  // 只要有填任何一科就回傳加總（不強制九科全填）
  return vals.length > 0 ? vals.reduce((a,b) => a + parseFloat(b), 0) : null;
}
function getFilledCount(sc) {
  return ACTIVE_SUBJECTS.map(s => sc && sc[s]).filter(v => v !== null && v !== undefined && v !== "").length;
}
function getAvg(sc) {
  const count = getFilledCount(sc);
  const total = getTotal(sc);
  return (count > 0 && total !== null) ? total / count : null;
}
function scoreKey(studentId, examId) { return `${studentId}_${examId}`; }
function getScores(studentId, examId) { return S.scores[scoreKey(studentId, examId)] || {}; }

// 分數格色
function scoreClass(v) {
  if (v === null || v === "") return "";
  v = parseFloat(v);
  if (v >= 80) return "score-high";
  if (v < 60)  return "score-low";
  return "";
}

// ── 今日日期 ──────────────────────────────────────────────────
(function() {
  const el = $("today-date"); if (!el) return;
  const d = new Date();
  const wd = ["星期日","星期一","星期二","星期三","星期四","星期五","星期六"][d.getDay()];
  el.textContent = d.getFullYear()+"年"+(d.getMonth()+1)+"月"+d.getDate()+"日 · "+wd;
})();

// 首頁不顯示任務列表

// ── 登入 ──────────────────────────────────────────────────────
async function doLogin() {
  const acc = $("lacc").value.trim();
  const pw  = $("lpw").value;
  hide("lerr"); hide("linfo");
  if (!acc || !pw) { $("lerr").textContent = "請輸入帳號與密碼"; show("lerr"); return; }

  $("login-btn").disabled = true;
  $("login-btn").textContent = "登入中...";

  try {
    let uid = null;
    if (db) {
      const email = (typeof ACCOUNT_EMAIL_MAP !== "undefined" && ACCOUNT_EMAIL_MAP[acc])
        ? ACCOUNT_EMAIL_MAP[acc]
        : acc + "@" + EMAIL_DOMAIN;
      const cred = await auth.signInWithEmailAndPassword(email, pw);
      uid = cred.user.uid;
    } else {
      if (acc !== ADMIN_ACCOUNT || pw !== ADMIN_PASSWORD) throw { code: "auth/wrong-password" };
      $("linfo").textContent = "⚠️ Firebase 未設定，目前為示範模式"; show("linfo");
      await new Promise(r => setTimeout(r, 600)); hide("linfo");
    }

    // ── 查詢 teachers 集合，取得班級設定 ──────────────────────
    let teacherDoc = null;
    const isSystemAdmin = acc === ADMIN_ACCOUNT;

    if (db && uid) {
      // 先用 uid 查，找不到再用帳號查
      const byUid = await db.collection("teachers").doc(uid).get();
      if (byUid.exists) {
        teacherDoc = byUid.data();
      } else {
        const byAcc = await db.collection("teachers").where("account", "==", acc).limit(1).get();
        if (!byAcc.empty) teacherDoc = byAcc.docs[0].data();
      }
    }

    if (teacherDoc) {
      // 使用 Firestore 班級設定
      PREFIX = teacherDoc.classPrefix + "_";
      CURRENT_CLASS.prefix      = teacherDoc.classPrefix;
      CURRENT_CLASS.className   = teacherDoc.className   || CLASS_NAME;
      CURRENT_CLASS.classYear   = teacherDoc.classYear   || CLASS_YEAR;
      CURRENT_CLASS.teacherName = teacherDoc.teacherName || acc;
      CURRENT_CLASS.isAdmin     = teacherDoc.isAdmin === true;

      // 套用自訂科目（若有設定）
      if (Array.isArray(teacherDoc.subjects) && teacherDoc.subjects.length > 0) {
        ACTIVE_SUBJECTS = teacherDoc.subjects;
      } else {
        ACTIVE_SUBJECTS = (typeof SUBJECTS !== "undefined") ? [...SUBJECTS] : [];
      }
      // 套用自訂段考（若有設定）
      if (Array.isArray(teacherDoc.exams) && teacherDoc.exams.length > 0) {
        ACTIVE_EXAMS = teacherDoc.exams;
      } else {
        ACTIVE_EXAMS = (typeof EXAMS !== "undefined") ? [...EXAMS] : [];
      }
    } else {
      // teachers 集合找不到文件 → 拒絕登入
      if (auth) auth.signOut().catch(()=>{});
      throw { code: "auth/unauthorized" };
    }

    // 管理者：載入所有班級清單供切換
    if (CURRENT_CLASS.isAdmin && db) {
      try {
        const allSnap = await db.collection("teachers").get();
        CURRENT_CLASS.allClasses = allSnap.docs.map(d => ({ id: d.id, ...d.data() }));
      } catch(e) { CURRENT_CLASS.allClasses = []; }
    }

    S.cur     = acc;
    S.isAdmin = CURRENT_CLASS.isAdmin;
    await loadAllData();
    showMainPage();

  } catch(err) {
    const msgs = {
      "auth/user-not-found":    "帳號不存在",
      "auth/wrong-password":    "密碼錯誤",
      "auth/invalid-credential":"帳號或密碼錯誤",
      "auth/too-many-requests": "嘗試次數過多，請稍後再試",
      "auth/invalid-api-key":   "Firebase 尚未設定，請更新 config.js",
      "auth/unauthorized":      "此帳號尚未授權，請聯絡管理員",
      "auth/invalid-email":     "Email 格式不正確",
      "auth/network-request-failed": "網路連線失敗，請檢查網路",
    };
    const msg = msgs[err.code] || ("帳號或密碼錯誤（" + (err.code || err.message || "未知錯誤") + "）");
    $("lerr").textContent = msg;
    show("lerr");
  } finally {
    $("login-btn").disabled = false;
    $("login-btn").textContent = "登入系統";
  }
}

document.addEventListener("keydown", e => {
  if (e.key === "Enter" && $("login-page") && !$("login-page").classList.contains("hidden")) doLogin();
});

function confirmLogout() { show("logout-modal-bg"); }
function doLogout() {
  if (auth) auth.signOut().catch(()=>{});
  S.cur = null; S.isAdmin = false; S.students = []; S.scores = {};
  S.teacherComments = {}; S.studentMemos = {};
  // 重置班級設定回 config.js 預設
  PREFIX = (typeof COLLECTION_PREFIX !== "undefined" ? COLLECTION_PREFIX : "grade") + "_";
  CURRENT_CLASS.prefix      = typeof COLLECTION_PREFIX !== "undefined" ? COLLECTION_PREFIX : "grade";
  CURRENT_CLASS.className   = CLASS_NAME;
  CURRENT_CLASS.classYear   = CLASS_YEAR;
  CURRENT_CLASS.teacherName = "";
  CURRENT_CLASS.isAdmin     = false;
  CURRENT_CLASS.allClasses  = [];
  hide("main-page"); show("login-page");
  $("lacc").value = ""; $("lpw").value = "";
  hide("logout-modal-bg");
  showToast("已成功登出 👋");
}

// ── 顯示主頁面 ────────────────────────────────────────────────
function showMainPage() {
  hide("login-page"); show("main-page");

  // 更新 header 顯示班級名稱
  const el = $("user-name");
  if (el) el.textContent = CURRENT_CLASS.className + "　" + (CURRENT_CLASS.teacherName || S.cur);

  // 更新左上角標題和瀏覽器 tab
  const titleEl = $("header-title");
  if (titleEl) titleEl.textContent = CURRENT_CLASS.className + " 成績追蹤系統";
  document.title = CURRENT_CLASS.className + " 成績追蹤系統";

  // 管理者：顯示班級切換下拉
  renderAdminClassSwitcher();

  buildExamOptions("overview-exam", "7上");
  buildExamOptions("input-exam", "7上");
  buildExamOptions("report-exam", "", true);
  switchPage("overview");
}

// ── 管理者班級切換 ────────────────────────────────────────────
function renderAdminClassSwitcher() {
  const wrap = $("admin-class-switcher");
  if (!wrap) return;
  if (!CURRENT_CLASS.isAdmin) { wrap.style.display = "none"; return; }

  // 依 classPrefix 去重，每個班級只顯示一次
  const seen = new Set();
  const uniqueClasses = CURRENT_CLASS.allClasses.filter(c => {
    if (!c.classPrefix || seen.has(c.classPrefix)) return false;
    seen.add(c.classPrefix);
    return true;
  }).sort((a,b) => (a.className||"").localeCompare(b.className||""));

  if (uniqueClasses.length <= 1) { wrap.style.display = "none"; return; }

  wrap.style.display = "flex";
  const options = uniqueClasses
    .map(c => `<option value="${c.classPrefix}" ${c.classPrefix===CURRENT_CLASS.prefix?"selected":""}>${c.className||c.classPrefix}</option>`)
    .join("");
  wrap.innerHTML = `
    <span style="font-size:12px;color:#C8A850;white-space:nowrap">切換班級：</span>
    <select onchange="switchClass(this.value)" style="font-size:12px;padding:3px 8px;border-radius:6px;border:1px solid #C8A850;background:#2A2618;color:#F5F0E8;cursor:pointer">
      ${options}
    </select>`;
}

async function switchClass(newPrefix) {
  if (newPrefix === CURRENT_CLASS.prefix) return;
  const cls = CURRENT_CLASS.allClasses.find(c => c.classPrefix === newPrefix);
  if (!cls) return;

  // 更新目前班級設定
  PREFIX = newPrefix + "_";
  CURRENT_CLASS.prefix    = newPrefix;
  CURRENT_CLASS.className = cls.className || newPrefix;
  CURRENT_CLASS.classYear = cls.classYear || CLASS_YEAR;

  // 套用切換後班級的自訂科目/段考
  if (Array.isArray(cls.subjects) && cls.subjects.length > 0) {
    ACTIVE_SUBJECTS = cls.subjects;
  } else {
    ACTIVE_SUBJECTS = (typeof SUBJECTS !== "undefined") ? [...SUBJECTS] : [];
  }
  if (Array.isArray(cls.exams) && cls.exams.length > 0) {
    ACTIVE_EXAMS = cls.exams;
  } else {
    ACTIVE_EXAMS = (typeof EXAMS !== "undefined") ? [...EXAMS] : [];
  }

  // 清除舊資料
  S.students = []; S.scores = {};
  S.teacherComments = {}; S.studentMemos = {};
  S.analysisStudentId = null; S.reportStudentId = null;
  S.examSubjectCountCache = {};

  const el = $("user-name");
  if (el) el.textContent = CURRENT_CLASS.className + "　（管理者）";
  const titleEl = $("header-title");
  if (titleEl) titleEl.textContent = CURRENT_CLASS.className + " 成績追蹤系統";
  document.title = CURRENT_CLASS.className + " 成績追蹤系統";

  showToast(`⏳ 切換至 ${CURRENT_CLASS.className}...`);
  await loadAllData();
  switchPage("overview");
  renderAdminClassSwitcher();
  showToast(`✅ 已切換至 ${CURRENT_CLASS.className}`);
}

// ── 頁面切換 ──────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════
// 裝置偵測
// ══════════════════════════════════════════════════════════
function isMobile(){ return window.innerWidth <= 768; }
function isIOS(){
  return /iPad|iPhone|iPod/.test(navigator.userAgent) ||
    (navigator.platform==='MacIntel' && navigator.maxTouchPoints>1);
}

// ══════════════════════════════════════════════════════════
// 裝置佈局切換（每次切換頁面或 resize 呼叫）
// ══════════════════════════════════════════════════════════
function applyDeviceLayout() {
  const mob = isMobile();
  // 總覽
  const mOvSel=$("m-ov-sel"), mOvW=$("m-ov-wrap"), pcOvW=$("pc-ov-wrap");
  if(mOvSel) mOvSel.style.display = mob?"block":"none";
  if(mOvW)   mOvW.style.display   = mob?"block":"none";
  if(pcOvW)  pcOvW.style.display  = mob?"none":"block";
  // 名單
  const mStW=$("m-st-wrap"), pcStW=$("pc-st-wrap");
  if(mStW)  mStW.style.display  = mob?"block":"none";
  if(pcStW) pcStW.style.display = mob?"none":"block";
  // 分析
  const mAnSel=$("m-an-sel"), mAnW=$("m-an-wrap"), pcAnW=$("pc-an-wrap");
  if(mAnSel) mAnSel.style.display = mob?"block":"none";
  if(mAnW)   mAnW.style.display   = mob?"block":"none";
  if(pcAnW)  pcAnW.style.display  = mob?"none":"block";
}
let resizeTimer;
function onViewportResize() {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => {
    applyDeviceLayout();
    if(isMobile()){
      if(S.activePage==="overview") renderMobileOverview($("m-overview-exam")?.value||$("overview-exam")?.value||"");
      if(S.activePage==="students") renderMobileStudentList();
      if(S.activePage==="analysis"&&S.analysisStudentId) renderMobileAnalysis();
    }
  }, 500);
}
// 優先用 visualViewport（iOS 滑動時不會誤觸），降級用 window resize
if (window.visualViewport) {
  window.visualViewport.addEventListener('resize', onViewportResize, {passive:true});
} else {
  window.addEventListener('resize', onViewportResize, {passive:true});
}

// ══════════════════════════════════════════════════════════
// 底部導覽列 tab 同步
// ══════════════════════════════════════════════════════════
function syncMobileNav(name) {
  ["overview","students","input","analysis","report"].forEach(p=>{
    const el=document.getElementById("mnav-"+p);
    if(el) el.classList.toggle("active", p===name);
  });
}

// ══════════════════════════════════════════════════════════
// 手機版總覽
// ══════════════════════════════════════════════════════════
function onMobileOverviewSemChange() {
  const sem = $("m-overview-sem")?.value||"7上";
  // 年級總結模式：不需要段考選單，直接渲染年級總結
  if (GRADE_OV_MAP[sem]) {
    const examSel = $("m-overview-exam");
    if (examSel) examSel.style.display = "none";
    renderMobileGradeOverview(GRADE_OV_MAP[sem]);
    return;
  }
  const examSel = $("m-overview-exam");
  if (examSel) examSel.style.display = "";
  buildExamOptions("m-overview-exam", sem);
  const examId = $("m-overview-exam")?.value;
  if(examId) renderMobileOverview(examId);
}

function renderMobileGradeOverview(gradeInfo) {
  const wrap = $("m-ov-wrap"); if (!wrap) return;
  const gradeExams = ACTIVE_EXAMS.filter(e => gradeInfo.semesters.includes(e.semester));
  // 各科跨段考平均
  const subAvgs = ACTIVE_SUBJECTS.map(sub => {
    const vals = gradeExams.flatMap(ex =>
      S.students.map(st => {
        const v = getScores(st.id, ex.id)[sub];
        return v !== undefined && v !== "" ? parseFloat(v) : null;
      })
    ).filter(v => v !== null);
    return { sub, avg: vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null, count: vals.length };
  }).filter(s => s.avg !== null);

  // 各次段考班平均總分
  const examAvgs = gradeExams.map(ex => {
    const ts = S.students.map(st=>getTotal(getScores(st.id,ex.id))).filter(v=>v!==null);
    return { ex, avg: ts.length?ts.reduce((a,b)=>a+b,0)/ts.length:null, n: ts.length };
  }).filter(e => e.avg !== null);

  const allTotals = gradeExams.flatMap(ex =>
    S.students.map(st => getTotal(getScores(st.id, ex.id))).filter(v=>v!==null)
  );
  const overallAvg = allTotals.length ? allTotals.reduce((a,b)=>a+b,0)/allTotals.length : null;

  let html = `
    <div class="m-hero">
      <div class="m-hero-lbl">${gradeInfo.label} · 各次段考平均</div>
      <div class="m-hero-val">${overallAvg !== null ? overallAvg.toFixed(1) : "—"}</div>
      <div class="m-hero-sub">共 ${gradeExams.length} 次段考 · ${S.students.length} 位學生</div>
    </div>`;

  // 各次段考趨勢
  if (examAvgs.length) {
    html += `<div class="m-sec">📈 各次段考班平均走勢</div>`;
    const maxAvg = Math.max(...examAvgs.map(e=>e.avg));
    examAvgs.forEach(({ex, avg: a}) => {
      const pct = maxAvg > 0 ? (a / maxAvg) * 100 : 0;
      const bc = a >= 60 ? "#2D5F8A" : "#A83232";
      html += `<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:11px">
        <div style="width:72px;text-align:right;color:#6B5F4A;flex-shrink:0">${ex.name.replace("次段考","").replace("第","")}</div>
        <div style="flex:1;height:14px;background:#F0EAE0;border-radius:3px;overflow:hidden">
          <div style="height:100%;width:${pct}%;background:${bc};border-radius:3px"></div>
        </div>
        <div style="width:36px;font-weight:700;font-family:monospace;font-size:12px;color:${bc}">${a.toFixed(1)}</div>
      </div>`;
    });
  }

  // 各科總平均
  html += `<div class="m-sec">📚 各科整體平均（${gradeInfo.label}）</div>`;
  subAvgs.forEach(({sub, avg: a}) => {
    const bc = a >= 80 ? "#5B8A4A" : a < 60 ? "#A83232" : "#2D5F8A";
    html += `<div class="m-sub-card">
      <div class="m-sub-name">${sub}</div>
      <div class="m-bar-row">
        <div class="m-bar"><div class="m-bar-fill" style="width:${a}%;background:${bc}"></div></div>
        <div class="m-bar-score" style="color:${bc}">${a.toFixed(1)}</div>
      </div>
    </div>`;
  });

  wrap.innerHTML = html;
}

function renderMobileOverview(examId) {
  const wrap=$("m-ov-wrap"); if(!wrap) return;
  if(!examId) examId=$("m-overview-exam")?.value||ACTIVE_EXAMS.find(e=>e.semester==="7上")?.id||"";
  const ex=ACTIVE_EXAMS.find(e=>e.id===examId); if(!ex) return;
  const totals=S.students.map(st=>getTotal(getScores(st.id,examId))).filter(v=>v!==null);
  const n=totals.length;
  const clsAvg=n?avg(totals):null;
  const highest=n?Math.max(...totals):null;
  const lowest=n?Math.min(...totals):null;
  const mMax=getExamMaxScore(examId);
  const passLine=mMax*0.6;
  const passRate=n?totals.filter(v=>v>=passLine).length/n*100:null;
  const COLORS=["#2D5F8A","#5B8A4A","#C4651A","#6B4FA0","#A83232"];
  const LIGHTS=["#EAF1F8","#EDF4EA","#FDF0E6","#F2EEF8","#FAECEC"];

  let html=`
    <div class="m-hero">
      <div class="m-hero-lbl">${ex.name} · 班級平均</div>
      <div class="m-hero-val">${clsAvg!==null?clsAvg.toFixed(1):"—"}</div>
      <div class="m-hero-sub">滿分 ${getExamMaxScore(examId)} 分 · ${n} 位有資料</div>
    </div>
    <div class="m-3col">
      <div class="m-box"><div class="mv" style="color:#2E5A1A">${highest!==null?highest.toFixed(0):"—"}</div><div class="ml">最高總分</div></div>
      <div class="m-box"><div class="mv" style="color:#8B2222">${lowest!==null?lowest.toFixed(0):"—"}</div><div class="ml">最低總分</div></div>
      <div class="m-box"><div class="mv" style="color:${passRate!==null?(passRate>=70?"#2E5A1A":passRate>=50?"#C4651A":"#8B2222"):"inherit"}">${passRate!==null?passRate.toFixed(0)+"%":"—"}</div><div class="ml">及格率</div></div>
    </div>
    <div class="m-sec">各科班平均</div>`;

  ACTIVE_SUBJECTS.forEach(sub=>{
    const vals=S.students.map(st=>{const v=getScores(st.id,examId)[sub];return v!==undefined&&v!==""?parseFloat(v):null;}).filter(v=>v!==null);
    const sa=vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null;
    const pc=vals.filter(v=>v>=60).length;
    const bc=sa===null?"#E0DAD0":sa>=80?"#5B8A4A":sa<60?"#A83232":"#2D5F8A";
    html+=`<div class="m-sub-card">
      <div class="m-sub-name">${sub}</div>
      <div class="m-bar-row">
        <div class="m-bar"><div class="m-bar-fill" style="width:${sa||0}%;background:${bc}"></div></div>
        <div class="m-bar-score" style="color:${bc}">${sa!==null?sa.toFixed(1):"—"}</div>
      </div>
      <div class="m-bar-meta">
        <span>及格 ${pc}/${vals.length} 人</span>
        <span>最高 ${vals.length?Math.max(...vals).toFixed(0):"—"} / 最低 ${vals.length?Math.min(...vals).toFixed(0):"—"}</span>
      </div>
    </div>`;
  });

  // 成績分佈（簡易長條）
  if (totals.length) {
    const brackets = [
      {label:"90+",  min:90,  max:101},
      {label:"80-89",min:80,  max:90},
      {label:"70-79",min:70,  max:80},
      {label:"60-69",min:60,  max:70},
      {label:"60以下",min:0,  max:60},
    ];
    const maxCount = Math.max(...brackets.map(b => totals.filter(t=>t>=b.min&&t<b.max).length), 1);
    html += `<div class="m-sec">📊 成績分佈</div>`;
    brackets.forEach(b => {
      const cnt = totals.filter(t => t >= b.min && t < b.max).length;
      const pct = (cnt / maxCount) * 100;
      const bc  = b.min >= 80 ? "#5B8A4A" : b.min >= 60 ? "#2D5F8A" : "#A83232";
      html += `<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:12px">
        <div style="width:52px;text-align:right;color:#6B5F4A;flex-shrink:0">${b.label}</div>
        <div style="flex:1;height:16px;background:#F0EAE0;border-radius:3px;overflow:hidden">
          <div style="height:100%;width:${pct}%;background:${bc};border-radius:3px;transition:width .3s"></div>
        </div>
        <div style="width:24px;font-weight:700;font-family:monospace;color:${bc}">${cnt}</div>
      </div>`;
    });
  }

  // 需關注
  const attn=S.students.map((st,i)=>{
    const sc=getScores(st.id,examId);
    const t=getTotal(sc);
    const fc=ACTIVE_SUBJECTS.filter(s=>sc[s]!==undefined&&sc[s]!==""&&parseFloat(sc[s])<60).length;
    return{st,i,t,fc};
  }).filter(x=>x.fc>=3||(x.t!==null&&x.t<passLine)).sort((a,b)=>(a.t||9999)-(b.t||9999)).slice(0,4);
  if(attn.length){
    html+=`<div class="m-sec">⚠️ 需關注</div>`;
    attn.forEach(({st,i,t,fc})=>{
      html+=`<div class="m-stu-card" onclick="jumpToAnalysisAndSwitch('${st.id}')">
        <div class="m-stu-av" style="background:${LIGHTS[i%5]};color:${COLORS[i%5]}">${st.name[0]}</div>
        <div class="m-stu-info"><div class="m-stu-name">${st.name}</div><div class="m-stu-meta">不及格 ${fc} 科</div></div>
        <div class="m-stu-right"><div class="m-stu-total" style="color:#8B2222">${t!==null?t.toFixed(0):"—"}</div></div>
      </div>`;
    });
  }
  wrap.innerHTML=html;
}

// ══════════════════════════════════════════════════════════
// 手機版名單
// ══════════════════════════════════════════════════════════
function renderMobileStudentList() {
  const wrap=$("m-st-wrap"); if(!wrap) return;
  if(!S.students.length){
    wrap.innerHTML='<div class="empty-state"><div class="empty-icon">👥</div><div class="empty-title">尚未新增學生</div></div>';
    return;
  }
  const COLORS=["#2D5F8A","#5B8A4A","#C4651A","#6B4FA0","#A83232"];
  const LIGHTS=["#EAF1F8","#EDF4EA","#FDF0E6","#F2EEF8","#FAECEC"];
  let html="";
  S.students.forEach((st,i)=>{
    const totals=ACTIVE_EXAMS.map(ex=>getTotal(getScores(st.id,ex.id)));
    const sv=totals.filter(v=>v!==null);
    const li=[...Array(ACTIVE_EXAMS.length).keys()].reverse().find(i=>totals[i]!==null);
    const pi=li>0?[...Array(li).keys()].reverse().find(i=>totals[i]!==null):undefined;
    const lt=li!==undefined?totals[li]:null;
    const pt=pi!==undefined?totals[pi]:null;
    const diff=(lt!==null&&pt!==null)?lt-pt:null;
    const lsc=li!==undefined?getScores(st.id,ACTIVE_EXAMS[li].id):{};
    const fc=ACTIVE_SUBJECTS.filter(s=>lsc[s]!==undefined&&lsc[s]!==""&&parseFloat(lsc[s])<60).length;
    const rank=lsc["班排"]||null;
    const stLastExId=li!==undefined?ACTIVE_EXAMS[li].id:null;
    const stMax=stLastExId?getStudentMaxScore(st.id,stLastExId):ACTIVE_SUBJECTS.length*100;
    const tc=lt===null?"#C8BA9E":lt>=stMax*0.7?"#2E5A1A":lt<stMax*0.6?"#8B2222":"#1C1A14";
    const spark=makeSpark(sv,56,20);
    const diffHtml=diff===null?""
      :diff>0?`<span class="m-pill m-pill-g">▲${diff.toFixed(0)}</span>`
      :diff<0?`<span class="m-pill m-pill-r">▼${Math.abs(diff).toFixed(0)}</span>`:"";
    let tag=fc>=3?`<span class="m-pill m-pill-r">待加強</span> `
      :lt!==null&&lt>=ACTIVE_SUBJECTS.length*80?`<span class="m-pill m-pill-g">優秀</span> `:"";
    html+=`<div class="m-stu-card" onclick="jumpToAnalysisAndSwitch('${st.id}')">
      <div class="m-stu-av" style="background:${LIGHTS[i%5]};color:${COLORS[i%5]}">${st.name[0]}</div>
      <div class="m-stu-info">
        <div class="m-stu-name">${String(st.number||"").padStart(2,"0")} ${st.name}</div>
        <div class="m-stu-meta">${tag}${rank?"班第"+rank+"名":""}</div>
      </div>
      <div class="m-stu-right">${spark}<div class="m-stu-total" style="color:${tc}">${lt!==null?lt.toFixed(0):"—"}</div><div class="m-stu-diff">${diffHtml}</div></div>
    </div>`;
  });
  wrap.innerHTML=html;
}

// ══════════════════════════════════════════════════════════
// 手機版分析
// ══════════════════════════════════════════════════════════
function populateMobileAnStudentSelect() {
  const el=$("m-an-student"); if(!el) return;
  const cur=el.value;
  el.innerHTML=`<option value="">— 選擇學生 —</option>`;
  S.students.forEach(st=>{el.innerHTML+=`<option value="${st.id}" ${cur===st.id?"selected":""}>${String(st.number||"").padStart(2,"0")} ${st.name}</option>`;});
  if(cur) el.value=cur;
}
function onMobileAnStudentChange() {
  S.analysisStudentId=$("m-an-student")?.value||"";
  const pcSel=$("analysis-student");
  if(pcSel) pcSel.value=S.analysisStudentId;
  renderMobileAnalysis();
}

function renderMobileGradeSummary(studentId, st, gradeInfo) {
  const wrap=$("m-an-wrap"); if(!wrap) return;
  const gradeExams  = ACTIVE_EXAMS.filter(e => gradeInfo.semesters.includes(e.semester));
  const gradeScores = gradeExams.map(ex => getScores(studentId, ex.id));
  const gradeTotals = gradeScores.map(sc => getTotal(sc));
  const validPairs  = gradeExams.map((ex,i)=>({ex, sc:gradeScores[i], t:gradeTotals[i]})).filter(x=>x.t!==null);

  if (!validPairs.length) {
    wrap.innerHTML=`<div class="empty-state"><div class="empty-icon">📊</div><div class="empty-title">此年級尚無資料</div></div>`;
    return;
  }

  const totalsArr = validPairs.map(p=>p.t);
  const avgTotal  = totalsArr.reduce((a,b)=>a+b,0)/totalsArr.length;
  const maxTotal  = Math.max(...totalsArr);
  const minTotal  = Math.min(...totalsArr);
  const firstT    = validPairs[0].t;
  const lastT     = validPairs[validPairs.length-1].t;
  const diff      = lastT - firstT;
  const lastSc    = validPairs[validPairs.length-1].sc;
  const filledSubs= ACTIVE_SUBJECTS.filter(s=>lastSc[s]!==undefined&&lastSc[s]!=="");
  const bestSub   = filledSubs.length?filledSubs.reduce((a,b)=>parseFloat(lastSc[b])>parseFloat(lastSc[a])?b:a):null;
  const failSubs  = filledSubs.filter(s=>parseFloat(lastSc[s])<60);
  const warnSubs  = ACTIVE_SUBJECTS.filter(sub=>{
    const vals=validPairs.slice(-2).map(p=>p.sc[sub]!==undefined&&p.sc[sub]!==""?parseFloat(p.sc[sub]):null).filter(v=>v!==null);
    return vals.length>=2&&vals.every(v=>v<60);
  });

  const subGradeAvgs=ACTIVE_SUBJECTS.map(sub=>{
    const vals=gradeScores.map(sc=>sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null).filter(v=>v!==null);
    return{sub,avg:vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null,count:vals.length};
  }).filter(s=>s.avg!==null);

  let lines=[];
  lines.push(diff>0?`📈 整個${gradeInfo.label}進步 <strong class="up">+${diff.toFixed(0)} 分</strong>`
    :diff<0?`📉 整個${gradeInfo.label}退步 <strong class="down">${diff.toFixed(0)} 分</strong>`
    :`➡️ ${gradeInfo.label}成績持平`);
  lines.push(`📊 年級平均總分 <strong>${avgTotal.toFixed(1)}</strong>（最高 ${maxTotal.toFixed(0)}、最低 ${minTotal.toFixed(0)}）`);
  if(bestSub)  lines.push(`💪 強項：<strong>${bestSub}</strong>（${parseFloat(lastSc[bestSub]).toFixed(0)} 分）`);
  if(failSubs.length) lines.push(`⚠️ 最近不及格：<strong class="down">${failSubs.join("、")}</strong>`);
  if(warnSubs.length) lines.push(`🚨 連續不及格：<strong class="down">${warnSubs.join("、")}</strong>`);

  let html=`
    <div class="m-analysis-top">
      <div class="m-analysis-name">${st.name}</div>
      <div class="m-analysis-meta">座號 ${st.number||"—"} ／ ${gradeInfo.label}</div>
    </div>
    <div class="m-summary">${lines.join("<br>")}</div>
    <div class="m-3col" style="margin-top:12px">
      <div class="m-box"><div class="mv">${validPairs.length}</div><div class="ml">有效段考</div></div>
      <div class="m-box"><div class="mv" style="color:#2E5A1A">${maxTotal.toFixed(0)}</div><div class="ml">最高總分</div></div>
      <div class="m-box"><div class="mv" style="color:#8B2222">${minTotal.toFixed(0)}</div><div class="ml">最低總分</div></div>
    </div>
    <div class="m-sec">各科年級平均</div>`;

  subGradeAvgs.forEach(({sub,avg,count})=>{
    const bc=avg>=80?"#5B8A4A":avg<60?"#A83232":"#2D5F8A";
    const lvl=avg>=90?"優秀":avg>=80?"良好":avg>=70?"中等":avg>=60?"及格":"待加強";
    html+=`<div class="m-sub-card">
      <div class="m-sub-name">${sub}</div>
      <div class="m-bar-row">
        <div class="m-bar"><div class="m-bar-fill" style="width:${avg}%;background:${bc}"></div></div>
        <div class="m-bar-score" style="color:${bc}">${avg.toFixed(1)}</div>
      </div>
      <div class="m-bar-meta"><span>${lvl}</span><span>${count} 次段考</span></div>
    </div>`;
  });

  html+=`<div class="m-sec">各次段考一覽</div>`;
  validPairs.forEach(({ex,sc,t})=>{
    const total=t!==null?t.toFixed(0):"—";
    const rank=sc["班排"]||"—";
    const examAvg=getAvg(sc);
    html+=`<div class="m-sub-card">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
        <span class="m-sub-name" style="margin:0;font-size:13px">${ex.name}</span>
        <span style="font-family:'DM Mono',monospace;font-weight:700;font-size:16px;color:#1C1A14">${total}</span>
      </div>
      <div class="m-bar-meta" style="margin-top:0">
        <span>班排 ${rank} 名</span>
        <span>平均 ${examAvg!==null?examAvg.toFixed(1):"—"} 分</span>
      </div>
    </div>`;
  });

  wrap.innerHTML=html;

  // 趨勢折線圖
  const chartContainer=document.createElement("div");
  chartContainer.className="m-sub-card";
  chartContainer.style.marginTop="8px";
  chartContainer.innerHTML=`<div class="m-sub-name">📈 ${gradeInfo.label}總分走勢</div><div style="height:200px;position:relative"><canvas id="m-chart-grade-trend"></canvas></div>`;
  wrap.appendChild(chartContainer);
  destroyChart("m-chart-grade-trend");
  chartInstances["m-chart-grade-trend"]=new Chart($("m-chart-grade-trend"),{
    type:"line",
    data:{labels:validPairs.map(p=>p.ex.name.replace("次段考","").replace("第","")),
      datasets:[{label:"總分",data:validPairs.map(p=>p.t),
        borderColor:"#2D5F8A",backgroundColor:"#2D5F8A18",
        borderWidth:2.5,pointRadius:5,fill:true,tension:0.3}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false}},
      scales:{y:{min:0,grid:{color:"#E2DED6"},ticks:{font:{size:10}}},
               x:{grid:{display:false},ticks:{font:{size:9}}}}}
  });

  // 雷達圖（個人 + 班級平均疊加）
  if(subGradeAvgs.length>=3){
    // 班級年級各科平均
    const mClassGradeAvgs = subGradeAvgs.map(({sub}) => {
      const vals = gradeExams.flatMap(ex =>
        S.students.map(st2 => {
          const v = getScores(st2.id, ex.id)[sub];
          return v !== undefined && v !== "" ? parseFloat(v) : null;
        })
      ).filter(v => v !== null);
      return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
    });
    const mHasClass = mClassGradeAvgs.some(v=>v!==null);

    const radarContainer=document.createElement("div");
    radarContainer.className="m-sub-card";
    radarContainer.style.marginTop="8px";
    radarContainer.innerHTML=`<div class="m-sub-name">🕸 科目雷達圖（${gradeInfo.label}各科平均）</div><div style="height:220px;position:relative"><canvas id="m-chart-grade-radar"></canvas></div>`;
    wrap.appendChild(radarContainer);
    destroyChart("m-chart-grade-radar");
    chartInstances["m-chart-grade-radar"]=new Chart($("m-chart-grade-radar"),{
      type:"radar",
      data:{labels:subGradeAvgs.map(s=>s.sub),datasets:[
        { label:st.name+" "+gradeInfo.label+"平均",
          data:subGradeAvgs.map(s=>s.avg),
          borderColor:"#6B4FA0",backgroundColor:"#6B4FA022",borderWidth:2.5,pointRadius:4 },
        ...(mHasClass?[{ label:"班級平均",
          data:mClassGradeAvgs,
          borderColor:"#C4651A",backgroundColor:"#C4651A0D",
          borderWidth:1.5,pointRadius:3,borderDash:[4,3] }]:[])
      ]},
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{display:mHasClass,position:"bottom",labels:{font:{size:10},boxWidth:10,padding:6}}},
        scales:{r:{min:0,max:100,ticks:{font:{size:10},stepSize:20},pointLabels:{font:{size:11}}}}}
    });
  }
}

function renderMobileAnalysis() {
  const wrap=$("m-an-wrap"); if(!wrap) return;
  const studentId=S.analysisStudentId;
  if(!studentId){
    wrap.innerHTML='<div class="empty-state"><div class="empty-icon">👤</div><div class="empty-title">請先選擇學生</div></div>';
    return;
  }
  const st=S.students.find(s=>s.id===studentId); if(!st) return;
  const sem=$("m-an-sem")?.value||"7上";

  // ── 年級總結模式 ──────────────────────────────────────────
  if(GRADE_MAP[sem]){ renderMobileGradeSummary(studentId, st, GRADE_MAP[sem]); return; }

  const semExams=ACTIVE_EXAMS.filter(e=>e.semester===sem);
  const semScores=semExams.map(ex=>getScores(studentId,ex.id));
  const semTotals=semScores.map(sc=>getTotal(sc));

  // 摘要
  const vp=semTotals.map((t,i)=>({t,sc:semScores[i],ex:semExams[i]})).filter(x=>x.t!==null);
  let sumHtml="";
  if(vp.length){
    const first=vp[0],last=vp[vp.length-1];
    const diff=vp.length>=2?last.t-first.t:null;
    const filled=ACTIVE_SUBJECTS.filter(s=>last.sc[s]!==undefined&&last.sc[s]!=="");
    const best=filled.length?filled.reduce((a,b)=>parseFloat(last.sc[b])>parseFloat(last.sc[a])?b:a):null;
    const fails=filled.filter(s=>parseFloat(last.sc[s])<60);
    const warnSubs=ACTIVE_SUBJECTS.filter(sub=>{
      const vals=vp.slice(-2).map(p=>p.sc[sub]!==undefined&&p.sc[sub]!==""?parseFloat(p.sc[sub]):null).filter(v=>v!==null);
      return vals.length>=2&&vals.every(v=>v<60);
    });
    let lines=[];
    if(diff!==null) lines.push(diff>0?`📈 整體進步 <strong class="up">+${diff.toFixed(0)} 分</strong>`:diff<0?`📉 整體退步 <strong class="down">${diff.toFixed(0)} 分</strong>`:"➡️ 成績持平");
    if(best) lines.push(`💪 強項：<strong>${best}</strong>（${parseFloat(last.sc[best]).toFixed(0)} 分）`);
    if(fails.length) lines.push(`⚠️ 不及格：<strong class="warn">${fails.join("、")}</strong>`);
    if(warnSubs.length) lines.push(`🚨 連續不及格：<strong class="down">${warnSubs.join("、")}</strong>`);
    sumHtml=`<div class="m-summary">${lines.join("<br>")}</div>`;
  }

  // chips — 記住目前選中的 idx，重渲時維持
  const prevActiveChip = document.querySelector(".m-chip.active");
  const prevChipIdx = prevActiveChip ? parseInt(prevActiveChip.id.replace("mchip-","")) : -1;

  let chipsHtml=`<div class="m-chips">`;
  semExams.forEach((ex,i)=>{
    chipsHtml+=`<div class="m-chip" id="mchip-${i}" onclick="selectMobileExam(${i},'${sem}')">${ex.name.replace(sem,"").trim()}</div>`;
  });
  chipsHtml+=`</div><div id="m-exam-detail"></div>`;

  wrap.innerHTML=`
    <div class="m-analysis-top">
      <div class="m-analysis-name">${st.name}</div>
      <div class="m-analysis-meta">座號 ${st.number||"—"} ／ ${sem}學期</div>
    </div>
    ${sumHtml}
    <div class="m-sec">各次段考</div>
    ${chipsHtml}`;

  // 維持原本選中的 idx；若無記錄則選第一筆有資料的
  const firstData=semTotals.findIndex(v=>v!==null);
  const targetIdx = (prevChipIdx >= 0 && prevChipIdx < semExams.length)
    ? prevChipIdx
    : (firstData >= 0 ? firstData : 0);
  selectMobileExam(targetIdx, sem);
}

function selectMobileExam(idx, sem) {
  if(!sem) sem=$("m-an-sem")?.value||"7上";
  document.querySelectorAll(".m-chip").forEach((c,i)=>c.classList.toggle("active",i===idx));
  const studentId=S.analysisStudentId; if(!studentId) return;
  const semExams=ACTIVE_EXAMS.filter(e=>e.semester===sem);
  const ex=semExams[idx]; if(!ex) return;
  const sc=getScores(studentId,ex.id);
  const total=getTotal(sc);
  const prevSc=idx>0?getScores(studentId,semExams[idx-1].id):null;
  const wrap=document.getElementById("m-exam-detail"); if(!wrap) return;

  if(getFilledCount(sc)===0){
    wrap.innerHTML=`<div class="empty-state small"><div class="empty-text">尚無此次段考資料</div></div>`;
    return;
  }
  let html=`<div class="m-hero" style="margin-top:0">
    <div class="m-hero-lbl">${ex.name} · 總分</div>
    <div class="m-hero-val">${total!==null?total.toFixed(0):"—"}</div>
    <div class="m-hero-sub">班排 ${sc["班排"]||"—"} 名　校排 ${sc["校排"]||"—"} 名　平均 ${getAvg(sc)!==null?getAvg(sc).toFixed(1):"—"} 分</div>
  </div>`;

  ACTIVE_SUBJECTS.forEach(sub=>{
    const val=sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null;
    const prev=prevSc&&prevSc[sub]!==undefined&&prevSc[sub]!==""?parseFloat(prevSc[sub]):null;
    const diff=(val!==null&&prev!==null)?val-prev:null;
    const subRank=val!==null?getSubjectRank(studentId,ex.id,sub):null;
    const subAvg=getExamSubjectAvg(ex.id,sub);
    const bc=val===null?"#E0DAD0":val>=80?"#5B8A4A":val<60?"#A83232":"#2D5F8A";
    const cf=val!==null&&val<60&&prev!==null&&prev<60;
    const diffHtml=diff===null?"":diff>0?`<span class="m-pill m-pill-g">▲${diff.toFixed(0)}</span>`:diff<0?`<span class="m-pill m-pill-r">▼${Math.abs(diff).toFixed(0)}</span>`:"";
    const level=val===null?"":val>=90?"優秀":val>=80?"良好":val>=70?"中等":val>=60?"及格":"待加強";
    html+=`<div class="m-sub-card" style="${cf?"border-color:#E8B8B8;background:#FFF8F8":""}">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
        <span class="m-sub-name" style="margin:0">${sub}</span>
        <div style="display:flex;gap:6px;align-items:center">${diffHtml}<span style="font-size:10px;color:#9E9890">${subRank?"班"+subRank+"名":""}</span></div>
      </div>
      <div class="m-bar-row">
        <div class="m-bar"><div class="m-bar-fill" style="width:${val||0}%;background:${bc}"></div></div>
        <div class="m-bar-score" style="color:${bc}">${val!==null?val.toFixed(0):"—"}</div>
      </div>
      <div class="m-bar-meta">
        <span>${level}</span>
        <span>班平均 ${subAvg!==null?subAvg.toFixed(1):"—"}${subAvg!==null&&val!==null?" ("+((val-subAvg)>=0?"+":"")+(val-subAvg).toFixed(1)+")":""}</span>
      </div>
    </div>`;
  });
  wrap.innerHTML=html;

  // ── 雷達圖（手機版，只顯示有填寫的科目）──────────────
  // 先在 html 裡加 canvas，再渲染圖表
  const radarContainer = document.createElement("div");
  radarContainer.className = "m-sub-card";
  radarContainer.style.marginTop = "8px";
  radarContainer.innerHTML = `
    <div class="m-sub-name">🕸 科目雷達圖</div>
    <div style="height:220px;position:relative"><canvas id="m-chart-radar"></canvas></div>`;
  wrap.appendChild(radarContainer);

  const mRadarSubs = ACTIVE_SUBJECTS.filter(s=>sc[s]!==undefined&&sc[s]!=="");
  const mRadarData = mRadarSubs.map(s=>parseFloat(sc[s]));
  destroyChart("m-chart-radar");
  if (mRadarSubs.length >= 3) {
    chartInstances["m-chart-radar"] = new Chart($("m-chart-radar"), {
      type:"radar",
      data:{ labels:mRadarSubs, datasets:[{
        label: ex.name,
        data: mRadarData,
        borderColor:"#6B4FA0", backgroundColor:"#6B4FA022", borderWidth:2, pointRadius:4
      }]},
      options:{ responsive:true, maintainAspectRatio:false,
        plugins:{legend:{display:false}},
        scales:{ r:{min:0, max:100, ticks:{font:{size:10},stepSize:20}, pointLabels:{font:{size:11}}} }
      }
    });
  }
}

function jumpToAnalysisAndSwitch(studentId) {
  S.analysisStudentId=studentId;
  const pcSel=$("analysis-student");
  if(pcSel) pcSel.value=studentId;
  const mSel=$("m-an-student");
  if(mSel) mSel.value=studentId;
  switchPage("analysis");
}

function switchPage(name) {
  ["overview","students","input","analysis","report","settings","summary","alert-settings","track"].forEach(p => {
    const el = $("page-"+p);
    if (el) el.classList.toggle("hidden", p !== name);
    const tab = $("tab-"+p);
    if (tab) tab.classList.toggle("active", p === name);
  });
  S.activePage = name;
  syncMobileNav(name);
  applyDeviceLayout();
  const mob = isMobile();
  if (name === "overview") {
    buildExamOptions("overview-exam", $("overview-sem")?.value||"7上");
    buildExamOptions("m-overview-exam", $("m-overview-sem")?.value||"7上");
    renderOverview();
    if(mob) renderMobileOverview($("m-overview-exam")?.value||$("overview-exam")?.value);
  }
  if (name === "students")  { renderStudentList(); if(mob) renderMobileStudentList(); }
  if (name === "input") {
    buildExamOptions("input-exam", $("input-sem")?.value||"7上");
    renderInputTable();
    const mwarn = $("mobile-input-warn");
    if (mwarn) mwarn.style.display = isMobile() ? "block" : "none";
  }
  if (name === "analysis")  {
    populateStudentSelects();
    populateMobileAnStudentSelect();
    if(S.analysisStudentId) { renderAnalysis(); if(mob) renderMobileAnalysis(); }
    else {
      const w=$("analysis-content");
      if(w) w.innerHTML='<div class="empty-state"><div class="empty-icon">👤</div><div class="empty-title">請先選擇學生</div></div>';
      if(mob) { const mw=$("m-an-wrap"); if(mw) mw.innerHTML='<div class="empty-state"><div class="empty-icon">👤</div><div class="empty-title">請先選擇學生</div></div>'; }
    }
  }
  if (name === "report")    { populateStudentSelects(); if(S.reportStudentId) renderReport(); }
  if (name === "settings")       initSettingsEditor();
  if (name === "summary")        renderSummaryPage();
  if (name === "alert-settings") { loadAlertConfig(); renderAlertList(); }
  if (name === "track") {
    // 初始化科目選單
    const sel = $("track-subject-sel");
    if (sel) {
      sel.innerHTML = ACTIVE_SUBJECTS.map(s => `<option value="${s}">${s}</option>`).join("");
    }
    renderSubjectTrackTable();
  }
}

// ── Firestore 資料載入 ────────────────────────────────────────
async function loadAllData() {
  if (!db) { loadLocalData(); return; }

  // 先清空舊資料，避免殘留上一個班級的資料
  S.students = [];
  S.scores   = {};
  S.teacherComments = {};
  S.studentMemos    = {};
  S.examSubjectCountCache = {};

  // 顯示 loading
  const loadingEl = $("app-loading");
  if (loadingEl) loadingEl.style.display = "flex";

  try {
    const stSnap = await col("students").orderBy("number").get();
    S.students = stSnap.docs.map(d => ({ id: d.id, ...d.data() }));

    const scSnap = await col("scores").get();
    S.scores = {};
    scSnap.docs.forEach(d => { S.scores[d.id] = d.data(); });

    // 載入評語
    try {
      const cmSnap = await col("config").doc("teacher_comments").get();
      if (cmSnap.exists) S.teacherComments = cmSnap.data();
    } catch(e) { /* 評語讀取失敗不影響主流程 */ }

    // 載入學生備忘錄
    try {
      const memoSnap = await col("config").doc("student_memos").get();
      if (memoSnap.exists) S.studentMemos = memoSnap.data();
    } catch(e) { /* 備忘錄讀取失敗不影響主流程 */ }

    showToast("✅ 資料已從雲端載入");
    refreshCurrentPage();
  } catch(e) {
    console.warn("Firestore 讀取失敗，嘗試本地備份:", e.message);
    loadLocalData();
    showToast("⚠️ 雲端連線失敗，使用本地資料");
    refreshCurrentPage();
  } finally {
    if (loadingEl) loadingEl.style.display = "none";
  }
}

// 資料載入完成後，重新渲染目前頁面
function refreshCurrentPage() {
  updateHeaderCount();
  if (S.activePage === "input") {
    buildExamOptions("input-exam", $("input-sem")?.value || "7上");
    renderInputTable();
  } else if (S.activePage === "overview") {
    buildExamOptions("overview-exam", $("overview-sem")?.value || "7上");
    renderOverview();
  } else if (S.activePage === "students") {
    renderStudentList();
    if(isMobile()) renderMobileStudentList();
  } else if (S.activePage === "analysis") {
    populateStudentSelects();
    if (S.analysisStudentId) renderAnalysis();
  } else if (S.activePage === "report") {
    populateStudentSelects();
    if (S.reportStudentId) renderReport();
  }
}

// 本地備份
function saveLocalData() {
  const p = CURRENT_CLASS.prefix || "grade";
  localStorage.setItem(`${p}-students`, JSON.stringify(S.students));
  localStorage.setItem(`${p}-scores`,   JSON.stringify(S.scores));
  localStorage.setItem(`${p}-memos`,    JSON.stringify(S.studentMemos));
}
function loadLocalData() {
  const p = CURRENT_CLASS.prefix || "grade";
  const st = localStorage.getItem(`${p}-students`) || localStorage.getItem("grade-113-3-students");
  const sc = localStorage.getItem(`${p}-scores`)   || localStorage.getItem("grade-113-3-scores");
  const cm = localStorage.getItem("grade-113-3-comments");
  const mm = localStorage.getItem(`${p}-memos`)    || localStorage.getItem("grade-113-3-memos");
  if (st) S.students = JSON.parse(st);
  if (sc) S.scores   = JSON.parse(sc);
  if (cm) S.teacherComments = JSON.parse(cm);
  if (mm) S.studentMemos    = JSON.parse(mm);
}

// ── 學生備忘錄儲存到 Firebase ─────────────────────────────────
async function saveStudentMemo(studentId, text) {
  S.studentMemos[studentId] = text;
  saveLocalData();
  if (!db) return;
  try {
    await col("config").doc("student_memos").set(S.studentMemos);
  } catch(e) {
    console.warn("備忘錄儲存失敗:", e.message);
  }
}
async function saveTeacherComments() {
  if (!db) {
    localStorage.setItem("grade-113-3-comments", JSON.stringify(S.teacherComments));
    return;
  }
  try {
    await col("config").doc("teacher_comments").set(S.teacherComments);
  } catch(e) {
    // 雲端失敗時備份到 localStorage
    localStorage.setItem("grade-113-3-comments", JSON.stringify(S.teacherComments));
    console.warn("評語儲存失敗:", e.message);
  }
}

// ── 學生 CRUD ─────────────────────────────────────────────────
function openAddStudent() {
  $("modal-student-title").textContent = "新增學生";
  $("s-name").value = ""; $("s-number").value = "";
  $("s-id-hidden").value = "";
  show("modal-student-bg");
}
function openEditStudent(id) {
  const st = S.students.find(s => s.id === id); if (!st) return;
  $("modal-student-title").textContent = "編輯學生";
  $("s-name").value   = st.name;
  $("s-number").value = st.number || "";
  $("s-id-hidden").value = id;
  show("modal-student-bg");
}
function closeStudentModal() { hide("modal-student-bg"); }

async function saveStudent() {
  const name   = $("s-name").value.trim();
  const number = parseInt($("s-number").value) || 0;
  const editId = $("s-id-hidden").value;
  if (!name) { showToast("請輸入學生姓名"); return; }
  if (!number) { showToast("請輸入座號"); return; }

  // 檢查座號是否與其他學生重複
  const duplicateNum = S.students.find(s => parseInt(s.number) === number && s.id !== editId);
  if (duplicateNum) { showToast(`⚠️ 座號 ${number} 已被「${duplicateNum.name}」使用，請重新輸入`); return; }

  try {
    if (editId) {
      if (db) await col("students").doc(editId).update({ name, number });
      const st = S.students.find(s => s.id === editId);
      if (st) { st.name = name; st.number = number; }
    } else {
      const newSt = { name, number, createdAt: new Date().toISOString() };
      if (db) {
        const ref = await col("students").add(newSt);
        S.students.push({ id: ref.id, ...newSt });
      } else {
        const id = Date.now().toString(36);
        S.students.push({ id, ...newSt });
      }
    }
    S.students.sort((a,b) => (a.number||999)-(b.number||999));
    saveLocalData();
    closeStudentModal();
    renderStudentList();
    updateHeaderCount();
    showToast(editId ? "✅ 學生資料已更新" : "✅ 學生已新增");
  } catch(e) { showToast("儲存失敗：" + e.message); }
}

async function deleteStudent(id) {
  const st = S.students.find(s => s.id === id);
  if (!st || !confirm(`確定要刪除「${st.name}」及其所有成績資料？`)) return;
  try {
    if (db) {
      await col("students").doc(id).delete();
      ACTIVE_EXAMS.forEach(async ex => {
        try { await col("scores").doc(scoreKey(id, ex.id)).delete(); } catch(e) {}
      });
    }
    S.students = S.students.filter(s => s.id !== id);
    ACTIVE_EXAMS.forEach(ex => { delete S.scores[scoreKey(id, ex.id)]; });
    saveLocalData();
    renderStudentList();
    updateHeaderCount();
    showToast("🗑️ 已刪除學生資料");
  } catch(e) { showToast("刪除失敗：" + e.message); }
}

function updateHeaderCount() {
  const el = $("student-count-header");
  if (el) el.textContent = `共 ${S.students.length} 位學生`;
}

// ── SVG Sparkline 工具 ───────────────────────────────────────
function makeSpark(values, w=80, h=28) {
  const pts = values.map((v,i) => v !== null ? [i, v] : null).filter(Boolean);
  if (pts.length < 2) return `<span style="font-size:10px;color:#C8BA9E">—</span>`;
  const xs = pts.map(p=>p[0]), ys = pts.map(p=>p[1]);
  const minX=Math.min(...xs), maxX=Math.max(...xs);
  const minY=Math.min(...ys), maxY=Math.max(...ys);
  const pad = 3;
  const scaleX = maxX===minX ? 1 : (w-pad*2)/(maxX-minX);
  const scaleY = maxY===minY ? 1 : (h-pad*2)/(maxY-minY);
  const coords = pts.map(([x,y]) => [
    pad + (x-minX)*scaleX,
    h - pad - (y-minY)*scaleY
  ]);
  const d = coords.map((c,i)=>(i===0?"M":"L")+c[0].toFixed(1)+","+c[1].toFixed(1)).join(" ");
  const lastY = coords[coords.length-1][1];
  const trend = pts.length>=2 ? pts[pts.length-1][1]-pts[0][1] : 0;
  const color = trend>0?"#2E5A1A":trend<0?"#8B2222":"#8B7355";
  // filled area
  const areaD = d + ` L${coords[coords.length-1][0].toFixed(1)},${h} L${coords[0][0].toFixed(1)},${h} Z`;
  return `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" style="overflow:visible">
    <path d="${areaD}" fill="${color}" fill-opacity="0.08"/>
    <path d="${d}" fill="none" stroke="${color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
    <circle cx="${coords[coords.length-1][0].toFixed(1)}" cy="${lastY.toFixed(1)}" r="2.5" fill="${color}"/>
  </svg>`;
}

// ── 學生名單頁 ────────────────────────────────────────────────
function renderStudentList() {
  // 確保桌機版容器可見
  applyDeviceLayout();
  const wrap = $("student-list-table");
  if (!S.students.length) {
    wrap.innerHTML = `<div class="empty-state"><div class="empty-icon">👥</div><div class="empty-title">尚未新增學生</div><div class="empty-text">點擊右上角「新增學生」開始建立名單</div></div>`;
    return;
  }

  // 計算每位學生的全學期總分陣列（依 ACTIVE_EXAMS 順序）
  const studentData = S.students.map(st => {
    const totals = ACTIVE_EXAMS.map(ex => getTotal(getScores(st.id, ex.id)));
    // 只取有資料的點，做為 Sparkline 數據
    const sparkValues = totals.filter(v => v !== null);
    const latestIdx = [...Array(ACTIVE_EXAMS.length).keys()].reverse().find(i => totals[i]!==null);
    const prevIdx   = latestIdx>0 ? [...Array(latestIdx).keys()].reverse().find(i=>totals[i]!==null) : undefined;
    const latestTotal = latestIdx!==undefined ? totals[latestIdx] : null;
    const prevTotal   = prevIdx!==undefined   ? totals[prevIdx]   : null;
    const diff = (latestTotal!==null&&prevTotal!==null) ? latestTotal-prevTotal : null;
    const latestSc = latestIdx!==undefined ? getScores(st.id, ACTIVE_EXAMS[latestIdx].id) : {};
    const latestRank = latestSc["班排"]||null;
    const latestSchoolRank = latestSc["校排"]||null;
    const failCount  = ACTIVE_SUBJECTS.filter(s=>latestSc[s]!==undefined&&latestSc[s]!==""&&parseFloat(latestSc[s])<60).length;
    // 連續退步判斷（最近三次有資料的）
    const consecutive = sparkValues.length>=3 &&
      sparkValues[sparkValues.length-1] < sparkValues[sparkValues.length-2] &&
      sparkValues[sparkValues.length-2] < sparkValues[sparkValues.length-3];
    return { st, sparkValues, latestTotal, prevTotal, diff, latestRank, failCount, consecutive, latestIdx };
  });

  let html = `<div class="table-wrap"><table>
    <thead><tr>
      <th style="width:52px">座號</th>
      <th>姓名</th>
      <th style="width:100px">趨勢</th>
      <th style="width:80px">最新總分</th>
      <th style="width:72px">進退步</th>
      <th style="width:60px">班排</th>
      <th style="width:60px">校排</th>
      <th style="width:100px">狀態</th>
      <th style="width:96px">操作</th>
    </tr></thead><tbody>`;

  studentData.forEach(({ st, sparkValues, latestTotal, diff, latestRank, latestSchoolRank, failCount, consecutive, latestIdx }) => {
    const spark = makeSpark(sparkValues);
    const diffHtml = diff===null ? `<span style="color:#C8BA9E;font-size:12px">—</span>`
      : diff>0 ? `<span class="rank-chip up">▲${diff.toFixed(0)}</span>`
      : diff<0 ? `<span class="rank-chip down">▼${Math.abs(diff).toFixed(0)}</span>`
      : `<span style="color:#9E9890;font-size:12px">持平</span>`;
    const stLastExId = latestIdx!==undefined ? ACTIVE_EXAMS[latestIdx].id : null;
    const stMax = stLastExId ? getStudentMaxScore(st.id, stLastExId) : ACTIVE_SUBJECTS.length*100;
    const totalColor = latestTotal===null?"#C8BA9E"
      : latestTotal>=stMax*0.7?"#2E5A1A"
      : latestTotal<stMax*0.6?"#8B2222":"#1C1A14";
    let tags = "";
    if (consecutive) tags += `<span class="stag stag-red">連退</span>`;
    if (failCount>=3) tags += `<span class="stag stag-amber">不及格${failCount}科</span>`;
    if (!tags && latestTotal!==null && stLastExId && latestTotal>=stMax*0.8) tags = `<span class="stag stag-green">優秀</span>`;

    html += `<tr>
      <td class="mono muted" style="text-align:center">${String(st.number||"").padStart(2,"0")||"—"}</td>
      <td class="bold" style="cursor:pointer;color:#2D5F8A;text-decoration:underline dotted" onclick="jumpToAnalysis('${st.id}')" title="點擊查看個人分析">${st.name}</td>
      <td style="padding:6px 10px">${spark}</td>
      <td style="text-align:center;font-family:'DM Mono',monospace;font-weight:700;font-size:15px;color:${totalColor}">${latestTotal!==null?latestTotal.toFixed(0):"—"}</td>
      <td style="text-align:center">${diffHtml}</td>
      <td style="text-align:center;font-size:13px;color:#6B5F4A">${latestRank||"—"}</td>
      <td style="text-align:center;font-size:13px;color:#6B4FA0;font-family:'DM Mono',monospace">${latestSchoolRank||"—"}</td>
      <td>${tags||'<span style="font-size:11px;color:#C8BA9E">—</span>'}</td>
      <td><div class="row-actions">
        <button class="btn btn-sm" onclick="openEditStudent('${st.id}')">編輯</button>
        <button class="btn btn-sm btn-danger" onclick="deleteStudent('${st.id}')">刪除</button>
      </div></td>
    </tr>`;
  });

  html += `</tbody></table></div>`;
  wrap.innerHTML = html;
}

// ── 快速跳轉到個人分析 ──────────────────────────────────
function jumpToAnalysis(studentId) {
  S.analysisStudentId = studentId;
  switchPage("analysis");
  const sel = $("analysis-student");
  if (sel) sel.value = studentId;
  renderAnalysis();
}

function renderInputTable() {
  const examId = $("input-exam").value;
  const wrap   = $("input-table-wrap");
  if (!S.students.length) {
    wrap.innerHTML = `<div class="alert alert-info">請先到「學生名單」頁面新增學生。</div>`;
    return;
  }
  let html = `<div class="table-wrap"><table><thead><tr><th>座號</th><th>姓名</th>`;
  ACTIVE_SUBJECTS.forEach(s => html += `<th style="min-width:66px">${s}</th>`);
  html += `<th style="min-width:56px">班排</th><th style="min-width:56px">校排</th><th style="min-width:64px">總分</th><th style="min-width:64px">平均</th></tr></thead><tbody>`;
  // 先計算目前的班排（供顯示用）
  calcClassRanks(examId);

  S.students.forEach((st, si) => {
    const sc = getScores(st.id, examId);
    html += `<tr><td class="mono muted">${String(st.number||"").padStart(2,"0")||"—"}</td><td class="bold nowrap">${st.name}</td>`;
    ACTIVE_SUBJECTS.forEach((sub, subi) => {
      const val = sc[sub] !== undefined ? sc[sub] : "";
      const subKey = sub.replace(/\s/g,"_");
      html += `<td><input type="number" min="0" max="100" step="0.5" value="${val}"
        id="inp_${st.id}_${subKey}"
        class="score-input"
        data-si="${si}" data-subi="${subi}"
        onchange="previewTotal('${st.id}', '${examId}'); markUnsaved()"
        onkeydown="scoreInputKeydown(event, ${si}, ${subi}, '${st.id}', '${examId}')"
      ></td>`;
    });
    const br = sc["班排"]||"—", cr = sc["校排"]||"";
    html += `<td style="text-align:center;font-family:'DM Mono',monospace;font-weight:600;color:#2D5F8A" id="rank_${st.id}">${br}</td>`;
    html += `<td><input type="number" min="1" value="${cr}" id="inp_${st.id}_校排" class="score-input"
      onchange="previewTotal('${st.id}', '${examId}'); markUnsaved()"
      onkeydown="if(event.key==='Enter'){event.preventDefault();this.blur();previewTotal('${st.id}','${examId}')}"
    ></td>`;
    const total = getTotal(sc);
    const avg_val = getAvg(sc);
    const filled = getFilledCount(sc);
    html += `<td class="score-cell bold" id="total_${st.id}">${total!==null?total.toFixed(0):"—"}${filled>0&&filled<ACTIVE_SUBJECTS.length?'<span style="font-size:10px;color:#9E9890;display:block">('+filled+'/'+ACTIVE_SUBJECTS.length+'科)</span>':''}</td>`;
    html += `<td class="score-cell" id="avg_${st.id}" style="color:#6B4FA0">${avg_val!==null?avg_val.toFixed(1):"—"}</td></tr>`;
  });
  // ── 底部統計列：各科班級平均 ────────────────────────────
  html += `<tfoot><tr style="background:#FAF7F0;border-top:2px solid #C8BA9E">
    <td class="mono muted" style="text-align:center">—</td>
    <td style="font-weight:600;font-size:12px;color:#6B5F4A">班級平均</td>`;
  ACTIVE_SUBJECTS.forEach(sub => {
    const subKey = sub.replace(/\s/g,"_");
    const vals = S.students.map(st2 => {
      const el2 = document.getElementById(`inp_${st2.id}_${subKey}`);
      return (el2 && el2.value !== "") ? parseFloat(el2.value) : (getScores(st2.id, examId)[sub] || null);
    }).filter(v => v !== null && !isNaN(v));
    const subAvg = vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
    const color = subAvg===null?"#C8BA9E":subAvg>=80?"#2E5A1A":subAvg<60?"#8B2222":"#2D5F8A";
    html += `<td style="text-align:center;font-family:'DM Mono',monospace;font-weight:600;font-size:13px;color:${color}" id="foot-avg-${subKey}">${subAvg!==null?subAvg.toFixed(1):"—"}</td>`;
  });
  html += `<td></td><td></td><td id="total-avg-row" style="text-align:center;font-family:'DM Mono',monospace;font-weight:700;color:#1C1A14"></td><td></td></tr></tfoot>`;
  html += `</tbody></table></div>`;
  // 計算總分平均
  const allTotals2 = S.students.map(st2 => getTotal(getScores(st2.id, examId))).filter(v=>v!==null);
  const totalAvg2 = allTotals2.length ? allTotals2.reduce((a,b)=>a+b,0)/allTotals2.length : null;
  wrap.innerHTML = html;
  const tavEl = $("total-avg-row");
  if (tavEl && totalAvg2!==null) tavEl.textContent = totalAvg2.toFixed(1);
  // 每次重新渲染輸入表格時，檢查是否有未還原的草稿
  setTimeout(checkAndRestoreDraft, 100);
}

function previewTotal(studentId, examId) {
  // 防呆：檢查剛剛被修改的輸入值是否在 0~100 範圍內
  ACTIVE_SUBJECTS.forEach(sub => {
    const el = document.getElementById(`inp_${studentId}_${sub.replace(/\s/g,"_")}`);
    if (!el || el.value === "") return;
    const v = parseFloat(el.value);
    if (isNaN(v) || v < 0 || v > 100) {
      el.style.borderColor = "#A83232";
      el.style.background  = "#FFF4F4";
      el.title = "分數需介於 0 ~ 100 之間";
    } else {
      el.style.borderColor = "";
      el.style.background  = "";
      el.title = "";
    }
  });
  const sc = {};
  ACTIVE_SUBJECTS.forEach(sub => {
    const el = document.getElementById(`inp_${studentId}_${sub.replace(/\s/g,"_")}`);
    if (el && el.value !== "") sc[sub] = parseFloat(el.value);
  });
  const total = getTotal(sc);
  const avg_val = getAvg(sc);
  const filled = getFilledCount(sc);
  const totalEl = $("total_"+studentId);
  if (totalEl) {
    totalEl.innerHTML = total !== null ? total.toFixed(0) : "—";
    if (filled > 0 && filled < ACTIVE_SUBJECTS.length) {
      totalEl.innerHTML += `<span style="font-size:10px;color:#9E9890;display:block">(${filled}/${ACTIVE_SUBJECTS.length}科)</span>`;
    }
  }
  const avgEl = $("avg_"+studentId);
  if (avgEl) avgEl.textContent = avg_val !== null ? avg_val.toFixed(1) : "—";

  // 即時更新底部班級平均列
  const examId2 = $("input-exam")?.value;
  if (examId2) {
    ACTIVE_SUBJECTS.forEach(sub => {
      const subKey = sub.replace(/\s/g,"_");
      const footId = "foot-avg-" + subKey;
      const footEl = document.getElementById(footId);
      if (!footEl) return;
      const vals = S.students.map(st2 => {
        const el2 = document.getElementById(`inp_${st2.id}_${subKey}`);
        return (el2 && el2.value !== "") ? parseFloat(el2.value) : (getScores(st2.id, examId2)[sub] != null ? parseFloat(getScores(st2.id, examId2)[sub]) : null);
      }).filter(v => v !== null && !isNaN(v));
      const subAvg = vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
      const color = subAvg===null?"#C8BA9E":subAvg>=80?"#2E5A1A":subAvg<60?"#8B2222":"#2D5F8A";
      footEl.textContent = subAvg !== null ? subAvg.toFixed(1) : "—";
      footEl.style.color = color;
    });
    // 總分班平均
    const allTotalsNow = S.students.map(st2 => {
      const sc2 = {};
      ACTIVE_SUBJECTS.forEach(sub => {
        const el2 = document.getElementById(`inp_${st2.id}_${sub.replace(/\s/g,"_")}`);
        if (el2 && el2.value !== "") sc2[sub] = parseFloat(el2.value);
        else if (getScores(st2.id, examId2)[sub] !== undefined) sc2[sub] = parseFloat(getScores(st2.id, examId2)[sub]);
      });
      return getTotal(sc2);
    }).filter(v=>v!==null);
    const tavEl = $("total-avg-row");
    if (tavEl) tavEl.textContent = allTotalsNow.length ? (allTotalsNow.reduce((a,b)=>a+b,0)/allTotalsNow.length).toFixed(1) : "—";
  // 即時重算班排
  const examId2 = $("input-exam")?.value;
    S.students.forEach(st2 => {
      const sc2 = {};
      ACTIVE_SUBJECTS.forEach(sub => {
        const el2 = document.getElementById(`inp_${st2.id}_${sub.replace(/\s/g,"_")}`);
        if (el2 && el2.value !== "") sc2[sub] = parseFloat(el2.value);
      });
      S.scores[scoreKey(st2.id, examId2)] = { ...(S.scores[scoreKey(st2.id, examId2)]||{}), ...sc2 };
    });
    calcClassRanks(examId2);
    // 更新畫面上所有班排顯示
    S.students.forEach(st2 => {
      const rankEl = $("rank_"+st2.id);
      if (rankEl) rankEl.textContent = S.scores[scoreKey(st2.id, examId2)]?.["班排"] || "—";
    });
  }
}

// ── 成績輸入鍵盤導航 ─────────────────────────────────────────
// Tab/Enter → 同列下一科；ArrowDown/Up → 同科下一/上一學生
function scoreInputKeydown(e, si, subi, studentId, examId) {
  const n = S.students.length;
  const m = ACTIVE_SUBJECTS.length;
  let nextSi = si, nextSubi = subi;

  if (e.key === "Tab" || e.key === "Enter") {
    e.preventDefault();
    if (e.shiftKey) {
      nextSubi = subi - 1;
      if (nextSubi < 0) { nextSubi = m - 1; nextSi = (si - 1 + n) % n; }
    } else {
      nextSubi = subi + 1;
      if (nextSubi >= m) { nextSubi = 0; nextSi = (si + 1) % n; }
    }
  } else if (e.key === "ArrowDown") {
    e.preventDefault(); nextSi = Math.min(si + 1, n - 1);
  } else if (e.key === "ArrowUp") {
    e.preventDefault(); nextSi = Math.max(si - 1, 0);
  } else { return; }

  const nextSt = S.students[nextSi];
  const nextSub = ACTIVE_SUBJECTS[nextSubi];
  const nextEl = document.getElementById(`inp_${nextSt.id}_${nextSub.replace(/\s/g,"_")}`);
  if (nextEl) { nextEl.focus(); nextEl.select(); }
}

// ── 未儲存提示 ───────────────────────────────────────────────
let _hasUnsaved = false;
let _autosaveTimer = null;
const AUTOSAVE_INTERVAL = 30000; // 30 秒

function markUnsaved() {
  if (_hasUnsaved) return;
  _hasUnsaved = true;
  const btn = $("save-scores-btn");
  if (btn) {
    btn.style.animation = "none";
    btn.style.background = "#C4651A";
    btn.textContent = "⚠️ 儲存成績";
  }
  const warn = $("unsaved-warn");
  if (warn) warn.style.display = "flex";
  // 啟動自動儲存草稿計時器
  startAutosave();
}

function clearUnsaved() {
  _hasUnsaved = false;
  const btn = $("save-scores-btn");
  if (btn) { btn.style.background = ""; btn.textContent = "💾 儲存全部成績"; }
  const warn = $("unsaved-warn");
  if (warn) warn.style.display = "none";
  stopAutosave();
  clearDraft();
}

// ── 自動儲存草稿 ─────────────────────────────────────────────
function getDraftKey() {
  const p = CURRENT_CLASS.prefix || COLLECTION_PREFIX;
  const examId = $("input-exam") ? $("input-exam").value : "unknown";
  return `${p}-draft-${examId}`;
}

function saveDraft() {
  if (!_hasUnsaved) return;
  const examId = $("input-exam") ? $("input-exam").value : null;
  if (!examId) return;

  const draft = {};
  S.students.forEach(st => {
    const sc = {};
    ACTIVE_SUBJECTS.forEach(sub => {
      const el = document.getElementById(`inp_${st.id}_${sub.replace(/\s/g,"_")}`);
      if (el && el.value !== "") sc[sub] = el.value;
    });
    const crEl = $(`inp_${st.id}_校排`);
    if (crEl && crEl.value !== "") sc["校排"] = crEl.value;
    if (Object.keys(sc).length > 0) draft[st.id] = sc;
  });

  try {
    localStorage.setItem(getDraftKey(), JSON.stringify({
      examId, savedAt: new Date().toISOString(), data: draft
    }));
    // 更新提示文字顯示最後草稿時間
    const warn = $("unsaved-warn");
    if (warn) {
      const t = new Date().toLocaleTimeString("zh-TW", { hour:"2-digit", minute:"2-digit", second:"2-digit" });
      warn.innerHTML = `⚠️ 有未儲存的成績&nbsp;<span style="font-weight:400;color:#8B5E2A">（草稿已於 ${t} 自動備份）</span>`;
    }
  } catch(e) { console.warn("草稿儲存失敗:", e); }
}

function loadDraft() {
  try {
    const raw = localStorage.getItem(getDraftKey());
    if (!raw) return null;
    return JSON.parse(raw);
  } catch(e) { return null; }
}

function clearDraft() {
  try { localStorage.removeItem(getDraftKey()); } catch(e) {}
}

function startAutosave() {
  stopAutosave();
  _autosaveTimer = setInterval(saveDraft, AUTOSAVE_INTERVAL);
}

function stopAutosave() {
  if (_autosaveTimer) { clearInterval(_autosaveTimer); _autosaveTimer = null; }
}

// ── 載入草稿（切換段考時呼叫）───────────────────────────────
function checkAndRestoreDraft() {
  const draft = loadDraft();
  if (!draft) return;
  const t = new Date(draft.savedAt).toLocaleTimeString("zh-TW", { hour:"2-digit", minute:"2-digit" });
  const restore = confirm(`📋 偵測到上次未儲存的草稿（備份時間：${t}）\n\n是否要還原草稿內容？\n（選「取消」將捨棄草稿）`);
  if (!restore) { clearDraft(); return; }

  // 還原草稿到輸入欄位
  Object.entries(draft.data).forEach(([stId, sc]) => {
    Object.entries(sc).forEach(([field, val]) => {
      if (field === "校排") {
        const el = $(`inp_${stId}_校排`);
        if (el) el.value = val;
      } else {
        const el = document.getElementById(`inp_${stId}_${field.replace(/\s/g,"_")}`);
        if (el) el.value = val;
      }
    });
  });
  markUnsaved();
  showToast("✅ 草稿已還原，請確認後按儲存");
}
window.addEventListener("beforeunload", e => {
  if (_hasUnsaved) { e.preventDefault(); e.returnValue = "有尚未儲存的成績，確定要離開嗎？"; }
});
function calcClassRanks(examId) {
  const pairs = S.students.map(st => ({
    id: st.id,
    total: getTotal(S.scores[scoreKey(st.id, examId)] || {})
  })).filter(p => p.total !== null);

  // 依總分降序排列
  pairs.sort((a,b) => b.total - a.total);

  // 同分同名次（dense rank）
  let rank = 1;
  pairs.forEach((p, i) => {
    if (i > 0 && p.total < pairs[i-1].total) rank = i + 1;
    const key = scoreKey(p.id, examId);
    if (S.scores[key]) S.scores[key]["班排"] = rank;
  });
}

async function saveAllScores() {
  const examId = $("input-exam").value;
  const btn = $("save-scores-btn");

  // 防呆：儲存前驗證所有分數
  let hasError = false;
  S.students.forEach(st => {
    ACTIVE_SUBJECTS.forEach(sub => {
      const el = document.getElementById(`inp_${st.id}_${sub.replace(/\s/g,"_")}`);
      if (!el || el.value === "") return;
      const v = parseFloat(el.value);
      if (isNaN(v) || v < 0 || v > 100) {
        el.style.borderColor = "#A83232";
        el.style.background  = "#FFF4F4";
        hasError = true;
      }
    });
  });
  if (hasError) { showToast("⚠️ 有分數超出範圍（0–100），請修正後再儲存"); return; }

  btn.disabled = true; btn.textContent = "儲存中...";

  // 步驟一：把目前畫面上這次段考的輸入值同步到 S.scores
  S.students.forEach(st => {
    const sc = {};
    ACTIVE_SUBJECTS.forEach(sub => {
      const el = document.getElementById(`inp_${st.id}_${sub.replace(/\s/g,"_")}`);
      if (el && el.value !== "") sc[sub] = parseFloat(el.value);
    });
    const crEl = $(`inp_${st.id}_校排`);
    if (crEl && crEl.value) sc["校排"] = parseInt(crEl.value);
    sc["_updatedAt"] = new Date().toISOString();
    S.scores[scoreKey(st.id, examId)] = sc;
  });

  // 步驟二：重算當前段考班排
  calcClassRanks(examId);

  // 步驟三：把「所有有資料的段考」全部寫入 Firestore（不只當前這次）
  const batch = db ? db.batch() : null;
  let savedCount = 0;

  ACTIVE_EXAMS.forEach(ex => {
    S.students.forEach(st => {
      const key = scoreKey(st.id, ex.id);
      const sc  = S.scores[key];
      if (!sc) return; // 沒有資料就跳過
      const hasData = ACTIVE_SUBJECTS.some(sub => sc[sub] !== undefined && sc[sub] !== "");
      if (!hasData) return; // 全空也跳過
      if (batch) batch.set(col("scores").doc(key), sc);
      savedCount++;
    });
  });

  try {
    if (batch) await batch.commit();
    saveLocalData();
    clearExamSubjectCountCache(); // 清除所有快取
    clearUnsaved();
    showToast(`✅ 已儲存全部 ${savedCount} 筆成績資料（含所有段考）`);
  } catch(e) {
    showToast("⚠️ 雲端儲存失敗，已存至本地：" + e.message);
    saveLocalData();
  } finally {
    btn.disabled = false; btn.textContent = "💾 儲存所有成績";
  }
}

// ── 班級總覽頁 ────────────────────────────────────────────────
let chartInstances = {};
function destroyChart(id) {
  if (chartInstances[id]) { chartInstances[id].destroy(); delete chartInstances[id]; }
}

const COLORS = ["#2D5F8A","#5B8A4A","#C4651A","#6B4FA0","#A83232","#1A7F7A","#B5860D","#5A3E8A","#2A7A5A"];

// ── 學期/段考選單連動 ────────────────────────────────────
// ── 班級總覽年級對應表 ───────────────────────────────────
const GRADE_OV_MAP = {
  "grade7":   { label:"7年級總結", semesters:["7上","7下"] },
  "grade8":   { label:"8年級總結", semesters:["8上","8下"] },
  "grade9":   { label:"9年級總結", semesters:["9上","9下"] },
  "gradeAll": { label:"全部總結",  semesters:["7上","7下","8上","8下","9上","9下"] },
};

function buildExamOptions(selectId, sem, includeAll=false, selectedVal="") {
  const el = $(selectId); if (!el) return;
  const filtered = sem ? ACTIVE_EXAMS.filter(e=>e.semester===sem) : ACTIVE_EXAMS;
  el.innerHTML = (includeAll ? `<option value="">全部段考</option>` : "") +
    filtered.map(e=>`<option value="${e.id}" ${e.id===selectedVal?"selected":""}>${e.name}</option>`).join("");
  // 確保有選到值（若 selectedVal 不在清單裡，預設選第一個）
  if (!includeAll && filtered.length > 0 && !filtered.find(e=>e.id===selectedVal)) {
    el.value = filtered[0].id;
  }
}
// ══════════════════════════════════════════════════════════
// 班級總覽：年級總結模式
// ══════════════════════════════════════════════════════════
function renderGradeOverview(gradeInfo) {
  const gradeExams = ACTIVE_EXAMS.filter(e => gradeInfo.semesters.includes(e.semester));
  const wrap = $("overview-stats");

  // 各次段考的班級平均總分
  const examAvgs = gradeExams.map(ex => {
    const totals = S.students.map(st=>getTotal(getScores(st.id,ex.id))).filter(v=>v!==null);
    return totals.length ? totals.reduce((a,b)=>a+b,0)/totals.length : null;
  });
  const validAvgs = examAvgs.filter(v=>v!==null);
  const overallAvg = validAvgs.length ? validAvgs.reduce((a,b)=>a+b,0)/validAvgs.length : null;
  const highestAvg = validAvgs.length ? Math.max(...validAvgs) : null;
  const lowestAvg  = validAvgs.length ? Math.min(...validAvgs) : null;
  const trend = validAvgs.length>=2 ? validAvgs[validAvgs.length-1]-validAvgs[0] : null;

  // 各科年級平均
  const subAvgs = ACTIVE_SUBJECTS.map(sub => {
    const vals = [];
    gradeExams.forEach(ex => {
      S.students.forEach(st => {
        const v = getScores(st.id,ex.id)[sub];
        if (v!==undefined&&v!=="") vals.push(parseFloat(v));
      });
    });
    return { sub, avg: vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null, count:vals.length };
  }).filter(s=>s.avg!==null);

  // 及格率（以各次段考平均計）
  const passLine = subAvgs.length*100*0.6;
  const passExams = gradeExams.map(ex=>{
    const totals=S.students.map(st=>getTotal(getScores(st.id,ex.id))).filter(v=>v!==null);
    const filled=totals.filter(v=>v>=passLine);
    return totals.length?filled.length/totals.length*100:null;
  }).filter(v=>v!==null);
  const avgPassRate = passExams.length ? passExams.reduce((a,b)=>a+b,0)/passExams.length : null;

  // 平均校排
  const schRanks = [];
  gradeExams.forEach(ex => {
    S.students.forEach(st => {
      const r = getScores(st.id,ex.id)["校排"];
      if (r) schRanks.push(parseInt(r));
    });
  });
  const avgSchRank = schRanks.length ? schRanks.reduce((a,b)=>a+b,0)/schRanks.length : null;

  // 統計卡片
  const trendHtml = trend===null?"" : trend>0
    ? `<span style="color:#2E5A1A;font-size:11px;margin-left:4px">▲${trend.toFixed(1)}</span>`
    : trend<0 ? `<span style="color:#8B2222;font-size:11px;margin-left:4px">▼${Math.abs(trend).toFixed(1)}</span>` : "";
  const cols = avgSchRank!==null ? 5 : 4;
  wrap.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;
  wrap.innerHTML = `
    <div class="stat-card">
      <div class="stat-label">${gradeInfo.label}</div>
      <div class="stat-value">${gradeExams.length}</div>
      <div class="stat-sub">次段考資料</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">年級平均總分</div>
      <div class="stat-value">${overallAvg!==null?overallAvg.toFixed(1):"—"}${trendHtml}</div>
      <div class="stat-sub">最高 ${highestAvg!==null?highestAvg.toFixed(1):"—"} / 最低 ${lowestAvg!==null?lowestAvg.toFixed(1):"—"}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">平均及格率</div>
      <div class="stat-value" style="color:${avgPassRate!==null?(avgPassRate>=70?"#2E5A1A":avgPassRate<50?"#8B2222":"#C4651A"):"inherit"}">${avgPassRate!==null?avgPassRate.toFixed(1)+"%":"—"}</div>
      <div class="stat-sub">各次段考平均</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">參與學生數</div>
      <div class="stat-value">${S.students.length}</div>
      <div class="stat-sub">位學生</div>
    </div>
    ${avgSchRank!==null?`<div class="stat-card"><div class="stat-label">平均校排</div><div class="stat-value" style="color:#C4651A">${avgSchRank.toFixed(1)}</div><div class="stat-sub">年級平均</div></div>`:""}
  `;

  // 各科年級平均表格
  const subWrap = $("subject-stats-wrap");
  if (subWrap) {
    let html = `<div style="font-size:12px;color:#6B5F4A;margin-bottom:8px">${gradeInfo.label}各科平均（跨 ${gradeExams.length} 次段考）</div>
      <div class="table-wrap"><table>
        <thead><tr><th>科目</th><th style="text-align:center">平均分</th><th style="min-width:140px">分布</th><th style="text-align:center">及格率</th><th style="text-align:center">級距</th></tr></thead>
        <tbody>`;
    subAvgs.forEach(({sub,avg,count})=>{
      const bc=avg>=80?"#5B8A4A":avg<60?"#A83232":"#2D5F8A";
      const lvl=avg>=90?"優秀":avg>=80?"良好":avg>=70?"中等":avg>=60?"及格":"待加強";
      // 及格率
      const vals2=[];
      gradeExams.forEach(ex=>{S.students.forEach(st=>{const v=getScores(st.id,ex.id)[sub];if(v!==undefined&&v!=="")vals2.push(parseFloat(v));});});
      const pr=vals2.length?vals2.filter(v=>v>=60).length/vals2.length*100:null;
      html+=`<tr>
        <td style="font-weight:600">${sub}</td>
        <td style="text-align:center;font-family:'DM Mono',monospace;font-weight:700;color:${bc}">${avg.toFixed(1)}</td>
        <td style="padding:6px 10px"><div style="height:8px;background:#F0EAE0;border-radius:4px;overflow:hidden"><div style="height:100%;width:${avg}%;background:${bc};border-radius:4px"></div></div></td>
        <td style="text-align:center;color:${pr!==null?(pr>=70?"#2E5A1A":pr<50?"#8B2222":"#C4651A"):"inherit"};font-weight:600">${pr!==null?pr.toFixed(1)+"%":"—"}</td>
        <td style="text-align:center"><span style="font-size:11px;padding:2px 8px;border-radius:99px;background:${bc}22;color:${bc};font-weight:600">${lvl}</span></td>
      </tr>`;
    });
    html+=`</tbody></table></div>`;
    subWrap.innerHTML=html;
  }

  // 趨勢折線圖（年級各次段考班平均）
  destroyChart("chart-avg-trend");
  const ctx=$("chart-avg-trend"); if(ctx) {
    chartInstances["chart-avg-trend"]=new Chart(ctx,{
      type:"line",
      data:{labels:gradeExams.map(e=>e.name.replace("次段考","").replace("第","")),
        datasets:[{label:"班級平均總分",data:examAvgs,
          borderColor:"#2D5F8A",backgroundColor:"#2D5F8A15",
          borderWidth:2.5,pointRadius:5,fill:true,tension:0.3,spanGaps:true}]},
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{display:false}},
        scales:{y:{grid:{color:"#E2DED6"},ticks:{font:{size:11}}},x:{grid:{display:false},ticks:{font:{size:9},maxRotation:45}}}
      }
    });
  }

  // 各科趨勢圖
  destroyChart("chart-subject-avg");
  const ctx2=$("chart-subject-avg"); if(ctx2) {
    chartInstances["chart-subject-avg"]=new Chart(ctx2,{
      type:"bar",
      data:{labels:subAvgs.map(s=>s.sub),
        datasets:[{label:"年級各科平均",
          data:subAvgs.map(s=>s.avg),
          backgroundColor:subAvgs.map(s=>s.avg>=80?"#5B8A4A99":s.avg<60?"#A8323299":"#2D5F8A99"),
          borderRadius:6}]},
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{display:false}},
        scales:{y:{min:0,max:100,grid:{color:"#E2DED6"},ticks:{font:{size:11}}},x:{grid:{display:false},ticks:{font:{size:11}}}}
      }
    });
  }

  // 清除不適用的區塊
  ["chart-dist","heatmap-wrap","school-rank-dist-wrap","class-strong-school-weak-wrap","attention-list","progress-list","subject-progress-wrap"].forEach(id=>{
    const el=$(id); if(el) el.innerHTML="";
  });

  // 年級弱科提醒 + 班級健康指數（依跨段考資料計算）
  renderClassWeakSubjects(null, gradeInfo);
  renderHealthIndex(null, gradeInfo);
}

function onOverviewSemChange() {
  const sem = $("overview-sem").value;
  const isGrade = !!GRADE_OV_MAP[sem];
  const examSel = $("overview-exam");
  if (examSel) examSel.style.display = isGrade ? "none" : "";
  if (!isGrade) buildExamOptions("overview-exam", sem);
  renderOverview();
}
function onInputSemChange() {
  const sem = $("input-sem").value;
  buildExamOptions("input-exam", sem);
  renderInputTable();
}
function onReportSemChange() {
  const sem = $("report-sem").value;
  buildExamOptions("report-exam", sem, true);
  renderReport();
}

function renderOverview() {
  const sem = $("overview-sem")?.value || "7上";
  const gradeInfo = GRADE_OV_MAP[sem];
  if (gradeInfo) { renderGradeOverview(gradeInfo); return; }
  const examId = $("overview-exam").value;
  renderOverviewStats(examId);
  renderAvgTrendChart();
  renderSubjectAvgChart(examId);
  renderClassWeakSubjects(examId);   // 項目 7：班級弱科警示
  renderDistChart(examId);
  renderHeatmap(examId);
  renderSchoolRankDist(examId);
  renderClassStrongSchoolWeak(examId);
  renderAttentionList(examId);
  renderProgressList();
  renderSubjectProgressSummary(examId);
  renderHealthIndex(examId);
  renderPrediction(examId);
}

// ── 班級弱科即時提醒 ─────────────────────────────────────────
function renderClassWeakSubjects(examId) {
  const wrap = $("class-weak-wrap"); if (!wrap) return;
  if (!examId) { wrap.innerHTML = ""; return; }

  // 計算每科：班平均、及格率、與校排平均的差距估算
  const subStats = ACTIVE_SUBJECTS.map(sub => {
    const vals = S.students.map(st => {
      const v = getScores(st.id, examId)[sub];
      return v!==undefined&&v!==""?parseFloat(v):null;
    }).filter(v=>v!==null);
    if (!vals.length) return null;
    const clsAvg   = vals.reduce((a,b)=>a+b,0)/vals.length;
    const passRate = vals.filter(v=>v>=60).length/vals.length*100;
    const sorted   = [...vals].sort((a,b)=>a-b);
    const n = sorted.length;
    const median = n%2===0?(sorted[n/2-1]+sorted[n/2])/2:sorted[Math.floor(n/2)];
    return { sub, clsAvg, passRate, median, n };
  }).filter(Boolean);

  if (!subStats.length) { wrap.innerHTML = ""; return; }

  // 找「弱科」：及格率 < 70% 或 班平均 < 65
  const weak = subStats
    .filter(s => s.passRate < 70 || s.clsAvg < 65)
    .sort((a,b) => a.clsAvg - b.clsAvg);

  // 找「強科」：班平均 >= 80
  const strong = subStats
    .filter(s => s.clsAvg >= 80)
    .sort((a,b) => b.clsAvg - a.clsAvg);

  if (!weak.length && !strong.length) {
    wrap.innerHTML = `<div style="font-size:12px;color:#9E9890;padding:8px">各科表現均衡，無明顯弱科 ✓</div>`;
    return;
  }

  let html = "";

  if (weak.length) {
    html += `<div style="margin-bottom:10px">
      <div style="font-size:12px;font-weight:700;color:#8B2222;margin-bottom:8px">⚠️ 需加強科目（${weak.length}科）</div>`;
    weak.forEach(({ sub, clsAvg, passRate, n }) => {
      const urgency = passRate < 50 ? "高" : passRate < 65 ? "中" : "低";
      const uColor  = urgency==="高"?"#8B2222":urgency==="中"?"#C4651A":"#8B7355";
      html += `<div style="display:flex;align-items:center;gap:10px;padding:8px 10px;background:#FFF8F8;border:1px solid #F4C4C4;border-radius:8px;margin-bottom:6px">
        <div style="min-width:44px;font-size:13px;font-weight:700;color:#8B2222">${sub}</div>
        <div style="flex:1">
          <div style="height:8px;background:#F0EAE0;border-radius:4px;overflow:hidden;margin-bottom:4px">
            <div style="height:100%;width:${clsAvg}%;background:${clsAvg<60?"#A83232":"#C4651A"};border-radius:4px"></div>
          </div>
          <div style="display:flex;gap:12px;font-size:11px;color:#6B5F4A">
            <span>班平均 <strong style="color:#8B2222">${clsAvg.toFixed(1)}</strong></span>
            <span>及格率 <strong style="color:#8B2222">${passRate.toFixed(0)}%</strong></span>
            <span>（${n} 人資料）</span>
          </div>
        </div>
        <span style="font-size:10px;padding:2px 7px;border-radius:99px;background:${uColor}22;color:${uColor};font-weight:700;white-space:nowrap">急迫度：${urgency}</span>
      </div>`;
    });
    html += `</div>`;
  }

  if (strong.length) {
    html += `<div>
      <div style="font-size:12px;font-weight:700;color:#2E5A1A;margin-bottom:8px">✨ 班級強項（${strong.length}科）</div>
      <div style="display:flex;flex-wrap:wrap;gap:8px">`;
    strong.forEach(({ sub, clsAvg, passRate }) => {
      html += `<div style="background:#EDF4EA;border:1px solid #B8DDB8;border-radius:8px;padding:6px 12px;font-size:12px">
        <span style="font-weight:700;color:#2E5A1A">${sub}</span>
        <span style="color:#6B5F4A;margin-left:6px">均${clsAvg.toFixed(1)} ／ 及格率${passRate.toFixed(0)}%</span>
      </div>`;
    });
    html += `</div></div>`;
  }

  wrap.innerHTML = html;
}

function renderDistChart(examId) {
  destroyChart("chart-dist");
  const ctx = $("chart-dist"); if (!ctx) return;
  const totals = S.students.map(st => getTotal(getScores(st.id, examId))).filter(v=>v!==null);
  if (!totals.length) return;

  const max = getExamMaxScore(examId);

  // 10 分一段的分佈
  const bucketSize = Math.ceil(max / 10);
  const buckets = [];
  for (let i = 0; i < 10; i++) {
    const lo = i * bucketSize, hi = (i+1) * bucketSize;
    const cnt = totals.filter(v => v >= lo && (i===9 ? v <= max : v < hi)).length;
    const pct  = lo / max;
    const color = pct>=0.8?"#2E5A1A":pct>=0.7?"#5B8A4A":pct>=0.6?"#2D5F8A":pct>=0.4?"#C4651A":"#A83232";
    buckets.push({ label:`${lo}–${hi>max?max:hi}`, cnt, color });
  }

  // 箱形圖五數摘要
  const sorted = [...totals].sort((a,b)=>a-b);
  const n = sorted.length;
  const q1 = sorted[Math.floor(n*0.25)];
  const q2 = n%2===0?(sorted[n/2-1]+sorted[n/2])/2:sorted[Math.floor(n/2)];
  const q3 = sorted[Math.floor(n*0.75)];
  const iqr = q3 - q1;
  const classAvgTotal = totals.reduce((a,b)=>a+b,0)/n;

  // 在 canvas 下方加五數摘要
  const statsId = "dist-five-num-" + examId;

  chartInstances["chart-dist"] = new Chart(ctx, {
    type:"bar",
    data:{ labels: buckets.map(b=>b.label),
      datasets:[{ label:"人數", data: buckets.map(b=>b.cnt),
        backgroundColor: buckets.map(b=>b.color+"CC"), borderRadius:5 }]},
    options:{ responsive:true, maintainAspectRatio:false, indexAxis:"y",
      plugins:{ legend:{display:false},
        tooltip:{ callbacks:{ label: ctx => ` ${ctx.parsed.x} 人（${(ctx.parsed.x/n*100).toFixed(0)}%）` } }
      },
      scales:{
        x:{ ticks:{stepSize:1,font:{size:11}}, grid:{color:"#E2DED6"},
            title:{display:true,text:"人數",font:{size:10},color:"#9E9890"} },
        y:{ grid:{display:false}, ticks:{font:{size:10}} }
      }
    }
  });

  // 五數摘要列
  const existingStats = document.getElementById(statsId);
  if (existingStats) existingStats.remove();
  const statsEl = document.createElement("div");
  statsEl.id = statsId;
  statsEl.style.cssText = "display:flex;gap:8px;flex-wrap:wrap;margin-top:8px;font-size:11px";
  [
    { label:"最低", val:Math.min(...totals).toFixed(0), color:"#A83232" },
    { label:"Q1",   val:q1.toFixed(0), color:"#C4651A" },
    { label:"中位數",val:q2.toFixed(1), color:"#2D5F8A" },
    { label:"平均", val:classAvgTotal.toFixed(1), color:"#6B4FA0" },
    { label:"Q3",   val:q3.toFixed(0), color:"#5B8A4A" },
    { label:"最高", val:Math.max(...totals).toFixed(0), color:"#2E5A1A" },
    { label:"IQR",  val:iqr.toFixed(0), color:"#6B5F4A" },
  ].forEach(({label,val,color})=>{
    statsEl.innerHTML += `<div style="flex:1;min-width:60px;text-align:center;background:#FAF7F0;border:1px solid #E0DAD0;border-radius:6px;padding:4px 6px">
      <div style="font-size:9px;color:#9E9890;letter-spacing:.05em">${label}</div>
      <div style="font-weight:700;font-family:'DM Mono',monospace;color:${color}">${val}</div>
    </div>`;
  });
  ctx.parentNode.after(statsEl);
}

function renderOverviewStats(examId) {
  // ── 上方 4 個總覽卡片（保留人數、平均、最高、最低）──────
  const wrap = $("overview-stats");
  const totals = S.students.map(st => getTotal(getScores(st.id, examId))).filter(v=>v!==null);
  const n = totals.length;
  const classAvg = n ? avg(totals) : null;
  const highest  = n ? Math.max(...totals) : null;
  const lowest   = n ? Math.min(...totals) : null;

  const avgSchRank = getExamAvgSchoolRank(examId);
  const minSchRank = (() => { const r=getExamSchoolRanks(examId).filter(v=>v!==null); return r.length?Math.min(...r):null; })();
  const cols = avgSchRank!==null ? 5 : 4;
  wrap.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;
  wrap.innerHTML = `
    <div class="stat-card"><div class="stat-label">班級人數</div><div class="stat-value">${S.students.length}</div><div class="stat-sub">有資料 ${n} 位</div></div>
    <div class="stat-card"><div class="stat-label">班級總分平均</div><div class="stat-value">${classAvg!==null?classAvg.toFixed(1):"—"}</div><div class="stat-sub">滿分 ${getExamMaxScore(examId)} 分</div></div>
    <div class="stat-card"><div class="stat-label">最高總分</div><div class="stat-value" style="color:#2E5A1A">${highest!==null?highest.toFixed(0):"—"}</div><div class="stat-sub">分</div></div>
    <div class="stat-card"><div class="stat-label">最低總分</div><div class="stat-value" style="color:#8B2222">${lowest!==null?lowest.toFixed(0):"—"}</div><div class="stat-sub">分</div></div>
    ${avgSchRank!==null?`<div class="stat-card"><div class="stat-label">平均校排</div><div class="stat-value" style="color:#C4651A">${avgSchRank.toFixed(1)}</div><div class="stat-sub">最佳 ${minSchRank} 名</div></div>`:""}
  `;

  // ── 下方：各科及格率 & 標準差表格 ──────────────────────
  renderSubjectStatsTable(examId);
}

function renderSubjectStatsTable(examId) {
  const wrap = $("subject-stats-wrap"); if (!wrap) return;
  if (!S.students.length) { wrap.innerHTML = ""; return; }

  const rows = ACTIVE_SUBJECTS.map(sub => {
    const vals = S.students
      .map(st => { const v = getScores(st.id, examId)[sub]; return v!==undefined&&v!==""?parseFloat(v):null; })
      .filter(v => v !== null);
    const n = vals.length;
    if (!n) return { sub, n:0, subAvg:null, median:null, passRate:null, stdDev:null, max:null, min:null };
    const subAvg   = vals.reduce((a,b)=>a+b,0) / n;
    const sorted   = [...vals].sort((a,b)=>a-b);
    const median   = n%2===0 ? (sorted[n/2-1]+sorted[n/2])/2 : sorted[Math.floor(n/2)];
    const passRate = vals.filter(v=>v>=60).length / n * 100;
    const stdDev   = n>1 ? Math.sqrt(vals.reduce((a,v)=>a+Math.pow(v-subAvg,2),0)/n) : 0;
    const max = Math.max(...vals), min = Math.min(...vals);
    return { sub, n, subAvg, median, passRate, stdDev, max, min };
  });

  const hasData = rows.some(r=>r.n>0);
  if (!hasData) {
    wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">尚無成績資料</div>';
    return;
  }

  let html = `<div class="table-wrap"><table>
    <thead><tr>
      <th style="width:60px">科目</th>
      <th style="width:52px;text-align:center">人數</th>
      <th style="width:64px;text-align:center">平均</th>
      <th style="width:64px;text-align:center" title="中位數不受極端值影響，更能反映班級中間水準">中位數</th>
      <th style="width:56px;text-align:center">最高</th>
      <th style="width:56px;text-align:center">最低</th>
      <th style="width:68px;text-align:center">及格率</th>
      <th style="min-width:140px">及格人數分布</th>
      <th style="width:64px;text-align:center">標準差</th>
      <th style="width:68px;text-align:center">分布狀況</th>
    </tr></thead><tbody>`;

  rows.forEach(({ sub, n, subAvg, median, passRate, stdDev, max, min }) => {
    if (!n) {
      html += `<tr><td style="font-weight:500">${sub}</td>
        <td colspan="9" style="text-align:center;color:#C8BA9E;font-size:12px">尚無資料</td></tr>`;
      return;
    }
    const passColor = passRate>=80?"#2E5A1A":passRate>=60?"#C4651A":"#8B2222";
    const passBar   = passRate>=80?"#5B8A4A":passRate>=60?"#C4651A":"#A83232";
    const stdColor  = stdDev>20?"#8B2222":stdDev<10?"#2E5A1A":"#C4651A";
    const stdLabel  = stdDev>20?"分散":stdDev<10?"集中":"適中";
    const avgColor  = subAvg>=80?"#2E5A1A":subAvg<60?"#8B2222":"#1C1A14";
    const passCount = Math.round(passRate/100*n);
    const failCount = n - passCount;
    // 平均 vs 中位數差距：若差距>5，說明有極端值偏移
    const medDiff   = subAvg - median;
    const medTitle  = Math.abs(medDiff)>5 ? (medDiff>0?"平均被高分拉高，多數人可能低於平均":"平均被低分拉低，多數人可能高於平均") : "平均與中位數接近，分布較對稱";
    const medColor  = Math.abs(medDiff)>5 ? "#C4651A" : "#2D5F8A";

    html += `<tr>
      <td style="font-weight:600">${sub}</td>
      <td style="text-align:center;color:#9E9890;font-size:12px">${n}</td>
      <td style="text-align:center;font-family:'DM Mono',monospace;font-weight:700;color:${avgColor}">${subAvg.toFixed(1)}</td>
      <td style="text-align:center;font-family:'DM Mono',monospace;font-weight:600;color:${medColor}" title="${medTitle}">${median.toFixed(1)}${Math.abs(medDiff)>5?`<span style="font-size:9px;display:block;color:${medColor}">${medDiff>0?"均↑":"均↓"}${Math.abs(medDiff).toFixed(1)}</span>`:""}</td>
      <td style="text-align:center;font-family:'DM Mono',monospace;color:#2E5A1A">${max.toFixed(0)}</td>
      <td style="text-align:center;font-family:'DM Mono',monospace;color:#8B2222">${min.toFixed(0)}</td>
      <td style="text-align:center;font-weight:700;color:${passColor}">${passRate.toFixed(1)}%</td>
      <td style="padding:6px 10px">
        <div style="display:flex;height:10px;border-radius:5px;overflow:hidden;background:#F0EAE0">
          <div style="width:${passRate.toFixed(1)}%;background:${passBar};transition:width .4s" title="及格 ${passCount} 人"></div>
          <div style="flex:1;background:#FAECEC" title="不及格 ${failCount} 人"></div>
        </div>
        <div style="display:flex;justify-content:space-between;font-size:10px;margin-top:3px">
          <span style="color:${passBar}">及格 ${passCount} 人</span>
          <span style="color:#A83232">不及格 ${failCount} 人</span>
        </div>
      </td>
      <td style="text-align:center;font-family:'DM Mono',monospace;font-weight:700;color:${stdColor}">${stdDev.toFixed(1)}</td>
      <td style="text-align:center">
        <span style="font-size:11px;padding:2px 8px;border-radius:99px;background:${stdColor}22;color:${stdColor};font-weight:600">${stdLabel}</span>
      </td>
    </tr>`;
  });

  html += `</tbody></table></div>`;
  wrap.innerHTML = html;
}

function getExamSubjectAvg(examId, subject) {
  const vals = S.students.map(st => {
    const sc = getScores(st.id, examId);
    return sc[subject] !== undefined && sc[subject] !== "" ? parseFloat(sc[subject]) : null;
  }).filter(v=>v!==null);
  return vals.length ? avg(vals) : null;
}

function renderAvgTrendChart() {
  destroyChart("chart-avg-trend");
  const ctx = $("chart-avg-trend"); if (!ctx) return;
  const sem = $("overview-sem")?.value || "7上";
  const semExams = ACTIVE_EXAMS.filter(e=>e.semester===sem);
  const datasets = ACTIVE_SUBJECTS.map((sub,i) => ({
    label: sub,
    data: semExams.map(ex => getExamSubjectAvg(ex.id, sub)),
    borderColor: COLORS[i], backgroundColor: COLORS[i]+"22",
    borderWidth: 2, pointRadius: 4, tension: 0.3,
    hidden: true
  }));
  chartInstances["chart-avg-trend"] = new Chart(ctx, {
    type: "line",
    data: { labels: semExams.map(e=>e.name), datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: "bottom", labels: { font:{size:11}, boxWidth:12, padding:8 }}},
      scales: {
        y: { min:0, max:100, grid:{color:"#E2DED6"}, ticks:{font:{size:11}}},
        x: { grid:{display:false}, ticks:{font:{size:11}}}
      }
    }
  });
}

function renderSubjectAvgChart(examId) {
  destroyChart("chart-subject-avg");
  const ctx = $("chart-subject-avg"); if (!ctx) return;
  const data = ACTIVE_SUBJECTS.map(sub => getExamSubjectAvg(examId, sub) || 0);
  chartInstances["chart-subject-avg"] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: ACTIVE_SUBJECTS,
      datasets: [{ label: "班級平均", data, backgroundColor: COLORS.map(c=>c+"CC"), borderRadius: 6 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: {display:false}},
      scales: {
        y: { min:0, max:100, grid:{color:"#E2DED6"}, ticks:{font:{size:11}}},
        x: { grid:{display:false}, ticks:{font:{size:11}}}
      }
    }
  });
}

function renderAttentionList(examId) {
  const wrap = $("attention-list");
  const items = S.students.map((st,i) => {
    const sc = getScores(st.id, examId);
    const total = getTotal(sc);
    const failCount = ACTIVE_SUBJECTS.filter(s => sc[s]!==undefined && sc[s]!=="" && parseFloat(sc[s])<60).length;
    return { st, i, total, failCount };
  }).filter(x => {
    if (x.failCount >= 3) return true;
    if (x.total === null) return false;
    return x.total < getStudentMaxScore(x.st.id, examId) * 0.55;
  });
  items.sort((a,b) => (a.total||9999)-(b.total||9999));

  if (!items.length) {
    wrap.innerHTML = `<div class="empty-state small"><div class="empty-text">本次段考無需特別關注的學生 🎉</div></div>`;
    return;
  }
  wrap.innerHTML = items.slice(0,6).map(({st,i,total,failCount}) => `
    <div class="attention-item">
      <div class="avatar" style="background:${["#EAF1F8","#EDF4EA","#FDF0E6","#F2EEF8","#FAECEC"][i%5]};color:${COLORS[i%COLORS.length]}">${st.name[0]}</div>
      <div class="attention-info"><div class="attention-name">${st.name}</div><div class="attention-sub">不及格科數：${failCount} 科</div></div>
      <span class="badge badge-amber">${total!==null?total.toFixed(0):"—"} 分</span>
    </div>`
  ).join("");
}

function renderProgressList() {
  const wrap = $("progress-list");
  const rows = S.students.map((st,i) => {
    const totals = ACTIVE_EXAMS.map(ex => getTotal(getScores(st.id, ex.id)));
    const available = totals.filter(v=>v!==null);
    if (available.length < 2) return null;
    const firstIdx = totals.findIndex(v=>v!==null);
    const lastIdx  = [...Array(totals.length).keys()].reverse().find(i=>totals[i]!==null);
    const first = totals[firstIdx], last = totals[lastIdx];
    // 找進步最多的科目
    const firstSc = getScores(st.id, ACTIVE_EXAMS[firstIdx].id);
    const lastSc  = getScores(st.id, ACTIVE_EXAMS[lastIdx].id);
    const subDiffs = ACTIVE_SUBJECTS.map(sub => {
      const f = parseFloat(firstSc[sub]), l = parseFloat(lastSc[sub]);
      return (!isNaN(f)&&!isNaN(l)) ? { sub, diff: l-f } : null;
    }).filter(Boolean);
    subDiffs.sort((a,b)=>b.diff-a.diff);
    const topSub = subDiffs[0];
    return { st, i, diff: last-first, last, topSub };
  }).filter(Boolean);
  rows.sort((a,b) => b.diff-a.diff);
  const top = rows.filter(p=>p.diff>0).slice(0,6);
  if (!top.length) {
    wrap.innerHTML = `<div class="empty-state small"><div class="empty-text">需至少兩次段考資料才能顯示</div></div>`;
    return;
  }
  wrap.innerHTML = top.map(({st,i,diff,last,topSub}) => `
    <div class="attention-item">
      <div class="avatar" style="background:${["#EAF1F8","#EDF4EA","#FDF0E6","#F2EEF8","#FAECEC"][i%5]};color:${COLORS[i%COLORS.length]}">${st.name[0]}</div>
      <div class="attention-info">
        <div class="attention-name">${st.name}</div>
        <div class="attention-sub">最新 ${last.toFixed(0)} 分${topSub&&topSub.diff>0?` · ${topSub.sub}進步最多(+${topSub.diff.toFixed(0)})`:""}</div>
      </div>
      <span class="badge badge-green">▲ ${diff.toFixed(0)}</span>
    </div>`
  ).join("");
}

// ── 科目維度：班級各科平均進退步（跨段考）──────────────────
function renderSubjectProgressSummary(examId) {
  const wrap = $("subject-progress-wrap"); if (!wrap) return;
  const sem = $("overview-sem")?.value || "7上";
  const semExams = ACTIVE_EXAMS.filter(e=>e.semester===sem);
  const curIdx = semExams.findIndex(e=>e.id===examId);
  if (curIdx <= 0 || semExams.length < 2) {
    wrap.innerHTML = `<div style="font-size:12px;color:#9E9890;padding:8px 0">選擇第二次以上的段考才能顯示科目進退步</div>`;
    return;
  }
  const prevExamId = semExams[curIdx-1].id;
  const rows = ACTIVE_SUBJECTS.map(sub => {
    const cur  = getExamSubjectAvg(examId, sub);
    const prev = getExamSubjectAvg(prevExamId, sub);
    const diff = (cur!==null&&prev!==null) ? cur-prev : null;
    return { sub, cur, prev, diff };
  });
  rows.sort((a,b) => (b.diff||0)-(a.diff||0));
  wrap.innerHTML = rows.map(({sub,cur,diff}) => {
    const color = diff===null?"#C8BA9E":diff>0?"#2E5A1A":diff<0?"#8B2222":"#9E9890";
    const arrow = diff===null?"—":diff>0?`▲${diff.toFixed(1)}`:diff<0?`▼${Math.abs(diff).toFixed(1)}`:"持平";
    return `<div style="display:flex;align-items:center;gap:8px;padding:4px 0;border-bottom:1px solid #F0EAE0">
      <span style="width:52px;font-size:12px;color:#6B5F4A;flex-shrink:0">${sub}</span>
      <div style="flex:1;height:6px;background:#F0EAE0;border-radius:3px;overflow:hidden">
        <div style="height:100%;width:${cur!==null?cur:0}%;background:${cur!==null?(cur>=80?"#5B8A4A":cur<60?"#A83232":"#2D5F8A"):"#E0DAD0"};border-radius:3px"></div>
      </div>
      <span style="width:32px;text-align:right;font-size:12px;font-family:'DM Mono',monospace;color:#1C1A14">${cur!==null?cur.toFixed(1):"—"}</span>
      <span style="width:44px;text-align:right;font-size:11px;font-weight:600;color:${color}">${arrow}</span>
    </div>`;
  }).join("");
}

// ── 計算某學生某科在班上的百分位（高於幾%的同學）──────────
function getSubjectPercentile(studentId, examId, subject) {
  const myVal = parseFloat(getScores(studentId, examId)[subject]);
  if (isNaN(myVal)) return null;
  const allVals = S.students
    .map(st => parseFloat(getScores(st.id, examId)[subject]))
    .filter(v => !isNaN(v));
  if (allVals.length < 2) return null;
  const below = allVals.filter(v => v < myVal).length;
  return Math.round(below / (allVals.length - 1) * 100);
}

// ── 計算某學生某科在班上的名次 ──────────────────────────────
function getSubjectRank(studentId, examId, subject) {
  const myVal = parseFloat(getScores(studentId, examId)[subject]);
  if (isNaN(myVal)) return null;
  const allVals = S.students
    .map(st => parseFloat(getScores(st.id, examId)[subject]))
    .filter(v => !isNaN(v));
  if (!allVals.length) return null;
  allVals.sort((a,b) => b-a);
  let rank = 1;
  for (let i=0; i<allVals.length; i++) {
    if (i>0 && allVals[i]<allVals[i-1]) rank = i+1;
    if (allVals[i] === myVal) return rank;
  }
  return null;
}


// ── 科目追蹤表（全班同一科目歷次趨勢）────────────────────────
function renderSubjectTrackTable() {
  const wrap = $("subject-track-wrap"); if (!wrap) return;
  const sub  = $("track-subject-sel")?.value;
  if (!sub || !S.students.length) {
    wrap.innerHTML = '<div class="empty-state small"><div class="empty-icon">📊</div><div class="empty-title">請選擇科目</div></div>';
    return;
  }

  const exams = ACTIVE_EXAMS;
  // 表格標頭
  let html = `<div class="table-wrap"><table>
    <thead><tr>
      <th style="min-width:32px">座號</th>
      <th style="min-width:64px">姓名</th>`;
  exams.forEach(ex => {
    html += `<th style="text-align:center;min-width:52px">${ex.name.replace("次段考","").replace("第","")}</th>`;
  });
  html += `<th style="text-align:center;min-width:52px">平均</th>
      <th style="text-align:center;min-width:60px">趨勢</th>
    </tr></thead><tbody>`;

  S.students.forEach(st => {
    const vals = exams.map(ex => {
      const v = getScores(st.id, ex.id)[sub];
      return (v !== undefined && v !== "") ? parseFloat(v) : null;
    });
    const filled = vals.filter(v => v !== null);
    const stAvg  = filled.length ? filled.reduce((a,b)=>a+b,0)/filled.length : null;
    const first  = filled.length ? vals.find(v=>v!==null) : null;
    const last   = filled.length ? [...vals].reverse().find(v=>v!==null) : null;
    const trend  = (first !== null && last !== null && filled.length >= 2) ? last - first : null;
    // 趨勢標籤（文字）
    const trendLabel = trend === null ? ""
      : trend > 5  ? `<div style="font-size:10px;color:#2E5A1A;font-weight:700;margin-top:2px">▲ +${trend.toFixed(0)}</div>`
      : trend < -5 ? `<div style="font-size:10px;color:#8B2222;font-weight:700;margin-top:2px">▼ ${trend.toFixed(0)}</div>`
      : `<div style="font-size:10px;color:#6B5F4A;margin-top:2px">≈ ${trend>0?"+":""}${trend.toFixed(0)}</div>`;
    // 趨勢折線（SVG sparkline，寬度稍大）
    const sparkHtml = makeSpark(vals, 72, 28);

    html += `<tr>
      <td class="mono">${String(st.number||"").padStart(2,"0")}</td>
      <td class="bold">${st.name}</td>`;
    vals.forEach(v => {
      const cls = v === null ? "muted" : v >= 80 ? "score-high" : v < 60 ? "score-low" : "";
      html += `<td class="score-cell ${cls}">${v !== null ? v.toFixed(0) : "—"}</td>`;
    });
    html += `<td class="score-cell bold" style="color:#1C4A6B">${stAvg !== null ? stAvg.toFixed(1) : "—"}</td>
      <td class="score-cell" style="min-width:80px">
        <div style="display:flex;flex-direction:column;align-items:center">${sparkHtml}${trendLabel}</div>
      </td>
    </tr>`;
  });

  // 班級平均列
  html += `<tr style="background:#FAF7F0;font-weight:700;border-top:2px solid #C8BA9E">
    <td colspan="2" style="color:#6B5F4A;font-size:12px;letter-spacing:.04em">班級平均</td>`;
  exams.forEach(ex => {
    const vs = S.students.map(st => {
      const v = getScores(st.id, ex.id)[sub];
      return (v !== undefined && v !== "") ? parseFloat(v) : null;
    }).filter(v => v !== null);
    const ea = vs.length ? vs.reduce((a,b)=>a+b,0)/vs.length : null;
    const cls = ea === null ? "muted" : ea >= 80 ? "score-high" : ea < 60 ? "score-low" : "";
    html += `<td class="score-cell ${cls}">${ea !== null ? ea.toFixed(1) : "—"}</td>`;
  });
  const allVals = ACTIVE_EXAMS.flatMap(ex =>
    S.students.map(st => { const v = getScores(st.id,ex.id)[sub]; return (v!==undefined&&v!=="") ? parseFloat(v) : null; })
  ).filter(v => v !== null);
  const allAvg = allVals.length ? allVals.reduce((a,b)=>a+b,0)/allVals.length : null;
  html += `<td class="score-cell score-high">${allAvg !== null ? allAvg.toFixed(1) : "—"}</td><td></td></tr>`;

  html += `</tbody></table></div>`;
  wrap.innerHTML = html;
}

// ── 個人分析頁 ────────────────────────────────────────────────
function populateStudentSelects() {
  ["analysis-student","report-student","parent-notice-student"].forEach(id => {
    const el = $(id); if (!el) return;
    const cur = el.value;
    el.innerHTML = `<option value="">— 選擇學生 —</option>`;
    S.students.forEach(st => {
      el.innerHTML += `<option value="${st.id}" ${cur===st.id?"selected":""}>${st.number?String(st.number).padStart(2,"0")+" ":"  "}${st.name}</option>`;
    });
    if (cur && S.students.find(s=>s.id===cur)) el.value = cur;
  });
}

function onAnalysisStudentChange() {
  S.analysisStudentId = $("analysis-student").value;
  // 切換學生後捲回頂部
  const pcWrap = $("pc-an-wrap");
  if (pcWrap) pcWrap.scrollTop = 0;
  const analysisPage = $("page-analysis");
  if (analysisPage) analysisPage.scrollTop = 0;
  window.scrollTo({top: 0, behavior: "smooth"});
  // 同步連動報告頁學生選單
  const repSel = $("report-student");
  if (repSel && S.analysisStudentId) {
    repSel.value = S.analysisStudentId;
    S.reportStudentId = S.analysisStudentId;
    const pns = $("parent-notice-student");
    if (pns) pns.value = S.analysisStudentId;
  }
  renderAnalysis();
}
function onAnalysisSemChange() {
  renderAnalysis();
}

// ── 年級總結輔助 ──────────────────────────────────────────────
const GRADE_MAP = {
  "grade7":   { label:"7年級總結", semesters:["7上","7下"] },
  "grade8":   { label:"8年級總結", semesters:["8上","8下"] },
  "grade9":   { label:"9年級總結", semesters:["9上","9下"] },
  "gradeAll": { label:"全部總結",  semesters:["7上","7下","8上","8下","9上","9下"] },
};

function renderGradeSummary(studentId, st, gradeInfo) {
  const wrap = $("analysis-content");
  const gradeExams  = ACTIVE_EXAMS.filter(e => gradeInfo.semesters.includes(e.semester));
  const gradeScores = gradeExams.map(ex => getScores(studentId, ex.id));
  const gradeTotals = gradeScores.map(sc => getTotal(sc));
  const validPairs  = gradeExams.map((ex,i) => ({ex, sc:gradeScores[i], t:gradeTotals[i]})).filter(x=>x.t!==null);

  if (!validPairs.length) {
    wrap.innerHTML = `<div class="empty-state"><div class="empty-icon">📊</div><div class="empty-title">此年級尚無資料</div></div>`;
    return;
  }

  const totalsArr = validPairs.map(p=>p.t);
  const avgTotal  = totalsArr.reduce((a,b)=>a+b,0)/totalsArr.length;
  const maxTotal  = Math.max(...totalsArr);
  const minTotal  = Math.min(...totalsArr);
  const firstT    = validPairs[0].t;
  const lastT     = validPairs[validPairs.length-1].t;
  const diff      = lastT - firstT;
  const lastSc    = validPairs[validPairs.length-1].sc;
  const filledSubs = ACTIVE_SUBJECTS.filter(s => lastSc[s]!==undefined && lastSc[s]!=="");
  const bestSub   = filledSubs.length ? filledSubs.reduce((a,b)=>parseFloat(lastSc[b])>parseFloat(lastSc[a])?b:a) : null;
  const failSubs  = filledSubs.filter(s => parseFloat(lastSc[s])<60);
  const warnSubs  = ACTIVE_SUBJECTS.filter(sub => {
    const vals = validPairs.slice(-2).map(p=>p.sc[sub]!==undefined&&p.sc[sub]!==""?parseFloat(p.sc[sub]):null).filter(v=>v!==null);
    return vals.length>=2 && vals.every(v=>v<60);
  });

  // 各科年級平均
  const subGradeAvgs = ACTIVE_SUBJECTS.map(sub => {
    const vals = gradeScores.map(sc=>sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null).filter(v=>v!==null);
    return { sub, avg: vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null, count:vals.length };
  }).filter(s=>s.avg!==null);

  // 摘要
  let lines = [];
  lines.push(diff>0 ? `📈 整個${gradeInfo.label}進步 <strong class="up">+${diff.toFixed(0)} 分</strong>`
    : diff<0 ? `📉 整個${gradeInfo.label}退步 <strong class="down">${diff.toFixed(0)} 分</strong>`
    : `➡️ ${gradeInfo.label}成績持平`);
  lines.push(`📊 年級平均總分 <strong>${avgTotal.toFixed(1)}</strong> 分（最高 ${maxTotal.toFixed(0)}、最低 ${minTotal.toFixed(0)}）`);
  if (bestSub) lines.push(`💪 強項科目：<strong>${bestSub}</strong>（${parseFloat(lastSc[bestSub]).toFixed(0)} 分）`);
  if (failSubs.length) lines.push(`⚠️ 最近不及格：<strong class="down">${failSubs.join("、")}</strong>`);
  if (warnSubs.length) lines.push(`🚨 連續不及格：<strong class="down">${warnSubs.join("、")}</strong>`);

  wrap.innerHTML = `
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;flex-wrap:wrap">
      <div style="font-size:20px;font-weight:700">${st.name}</div>
      <span style="color:#9E9890;font-size:13px">座號 ${st.number||"—"}</span>
      <span style="background:#1C1A14;color:#F5F0E8;font-size:12px;padding:3px 12px;border-radius:99px;font-weight:600">${gradeInfo.label}</span>
    </div>
    <div class="summary-box mb-16">
      <div class="summary-title">${gradeInfo.label} 學習摘要</div>
      <div class="summary-text">${lines.join("<br>")}</div>
    </div>
    <div class="grid4 mb-16">
      <div class="stat-card"><div class="stat-label">有效段考數</div><div class="stat-value">${validPairs.length}</div><div class="stat-sub">共 ${gradeExams.length} 次</div></div>
      <div class="stat-card"><div class="stat-label">年級平均總分</div><div class="stat-value">${avgTotal.toFixed(1)}</div><div class="stat-sub">分</div></div>
      <div class="stat-card"><div class="stat-label">最高總分</div><div class="stat-value" style="color:#2E5A1A">${maxTotal.toFixed(0)}</div><div class="stat-sub">${validPairs.find(p=>p.t===maxTotal)?.ex.name||""}</div></div>
      <div class="stat-card"><div class="stat-label">最低總分</div><div class="stat-value" style="color:#8B2222">${minTotal.toFixed(0)}</div><div class="stat-sub">${validPairs.find(p=>p.t===minTotal)?.ex.name||""}</div></div>
    </div>
    <div class="grid2 mb-16">
      <div class="card"><div class="card-title">📈 ${gradeInfo.label}總分走勢</div>
        <div class="chart-box" style="height:220px"><canvas id="chart-grade-trend"></canvas></div>
      </div>
      <div class="card"><div class="card-title">🕸 科目雷達圖（${gradeInfo.label}各科平均）</div>
        <div class="chart-box" style="height:220px"><canvas id="chart-grade-radar"></canvas></div>
      </div>
    </div>
    <div class="card mb-16">
      <div class="card-title">📊 ${gradeInfo.label}各科平均</div>
      ${subGradeAvgs.map(({sub,avg,count})=>{
        const bc=avg>=80?"#5B8A4A":avg<60?"#A83232":"#2D5F8A";
        const lvl=avg>=90?"優秀":avg>=80?"良好":avg>=70?"中等":avg>=60?"及格":"待加強";
        return `<div style="display:flex;align-items:center;gap:10px;padding:7px 0;border-bottom:1px solid #F0EAE0">
          <span style="width:56px;font-size:13px;font-weight:600">${sub}</span>
          <div style="flex:1;height:8px;background:#F0EAE0;border-radius:4px;overflow:hidden">
            <div style="height:100%;width:${avg}%;background:${bc};border-radius:4px"></div>
          </div>
          <span style="font-size:14px;font-weight:700;font-family:'DM Mono',monospace;color:${bc};min-width:36px;text-align:right">${avg.toFixed(1)}</span>
          <span style="font-size:10px;padding:1px 7px;border-radius:99px;background:${bc}22;color:${bc};font-weight:600">${lvl}</span>
          <span style="font-size:10px;color:#9E9890;min-width:36px">${count}次</span>
        </div>`;
      }).join("")}
    </div>
    <div class="card">
      <div class="card-title">🗂 ${gradeInfo.label}各次段考一覽</div>
      <div class="table-wrap"><table>
        <thead><tr><th>段考</th>${ACTIVE_SUBJECTS.map(s=>`<th style="text-align:center;min-width:44px">${s}</th>`).join("")}<th style="text-align:center">總分</th><th style="text-align:center">班排</th><th style="text-align:center">校排</th></tr></thead>
        <tbody>${validPairs.map(({ex,sc,t})=>`<tr>
          <td style="white-space:nowrap;font-size:12px">${ex.name}</td>
          ${ACTIVE_SUBJECTS.map(sub=>{const v=sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null;const c=v===null?"#C8BA9E":v>=80?"#2E5A1A":v<60?"#8B2222":"#1C1A14";return `<td style="text-align:center;font-family:'DM Mono',monospace;font-size:12px;font-weight:600;color:${c}">${v!==null?v.toFixed(0):"—"}</td>`;}).join("")}
          <td style="text-align:center;font-weight:700;font-family:'DM Mono',monospace">${t!==null?t.toFixed(0):"—"}</td>
          <td style="text-align:center;color:#6B5F4A;font-size:12px">${sc["班排"]||"—"}</td>
          <td style="text-align:center;color:#6B4FA0;font-size:12px;font-family:'DM Mono',monospace">${sc["校排"]||"—"}</td>
        </tr>`).join("")}</tbody>
      </table></div>
    </div>`;

  // 折線圖
  // 年級總結走勢 + 校排第二 Y 軸
  const gradeTrendLabels = validPairs.map(p=>p.ex.name.replace("次段考","").replace("第",""));
  const gradeTotalData   = validPairs.map(p=>p.t);
  const gradeRankData    = validPairs.map(p => p.sc["校排"] ? parseInt(p.sc["校排"]) : null);
  const hasGradeRank     = gradeRankData.some(v=>v!==null);
  const gradeRankMax     = hasGradeRank ? Math.max(...gradeRankData.filter(v=>v!==null)) + 5 : 50;

  destroyChart("chart-grade-trend");
  chartInstances["chart-grade-trend"] = new Chart($("chart-grade-trend"), {
    type:"line",
    data:{ labels: gradeTrendLabels,
      datasets:[
        { label:"總分", data:gradeTotalData,
          borderColor:"#2D5F8A", backgroundColor:"#2D5F8A18",
          borderWidth:2.5, pointRadius:5, fill:true, tension:0.3, yAxisID:"yTotal" },
        ...(hasGradeRank ? [{ label:"校排", data:gradeRankData,
          borderColor:"#C4651A", backgroundColor:"transparent",
          borderWidth:1.5, pointRadius:4, borderDash:[5,3],
          tension:0.3, spanGaps:true, yAxisID:"yRank" }] : [])
      ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ display:hasGradeRank, position:"bottom", labels:{font:{size:10},boxWidth:10,padding:8} } },
      scales:{
        yTotal:{ type:"linear", position:"left", grid:{color:"#E2DED6"}, ticks:{font:{size:11}},
          title:{display:true,text:"總分",font:{size:10},color:"#2D5F8A"} },
        ...(hasGradeRank ? { yRank:{ type:"linear", position:"right", reverse:true,
          min:1, max:gradeRankMax,
          grid:{drawOnChartArea:false}, ticks:{font:{size:10},color:"#C4651A"},
          title:{display:true,text:"校排",font:{size:10},color:"#C4651A"} } } : {})
      }
    }
  });

  // 雷達圖（個人年級平均 + 班級年級平均疊加）
  const rSubs = subGradeAvgs.map(s=>s.sub);
  const rData = subGradeAvgs.map(s=>s.avg);

  // 班級年級各科平均
  const classGradeAvgs = rSubs.map(sub => {
    const vals = gradeExams.flatMap(ex =>
      S.students.map(st2 => {
        const v = getScores(st2.id, ex.id)[sub];
        return v !== undefined && v !== "" ? parseFloat(v) : null;
      })
    ).filter(v => v !== null);
    return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
  });
  const hasClassGrade = classGradeAvgs.some(v=>v!==null);

  destroyChart("chart-grade-radar");
  if (rSubs.length) {
    chartInstances["chart-grade-radar"] = new Chart($("chart-grade-radar"), {
      type:"radar",
      data:{ labels:rSubs, datasets:[
        { label: st.name + " " + gradeInfo.label + "平均",
          data:rData, borderColor:"#6B4FA0", backgroundColor:"#6B4FA022", borderWidth:2.5, pointRadius:4 },
        ...(hasClassGrade ? [{ label:"班級平均",
          data: classGradeAvgs,
          borderColor:"#C4651A", backgroundColor:"#C4651A0D",
          borderWidth:1.5, pointRadius:3, borderDash:[4,3] }] : [])
      ]},
      options:{ responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{ display:hasClassGrade, position:"bottom", labels:{font:{size:10},boxWidth:10,padding:6} } },
        scales:{ r:{min:0, max:100, ticks:{font:{size:10},stepSize:20}, pointLabels:{font:{size:11}}} }
      }
    });
  }
}

function renderAnalysis() {
  const studentId = S.analysisStudentId;
  const sem = $("analysis-sem")?.value || "7上";
  const wrap = $("analysis-content");

  if (!studentId) {
    wrap.innerHTML = `<div class="empty-state"><div class="empty-icon">👤</div><div class="empty-title">請先選擇學生</div><div class="empty-text">從上方選擇學生與學期後開始分析</div></div>`;
    return;
  }
  const st = S.students.find(s=>s.id===studentId); if (!st) return;

  // 年級總結模式
  if (GRADE_MAP[sem]) { renderGradeSummary(studentId, st, GRADE_MAP[sem]); return; }

  // 取本學期三次段考
  const semExams  = ACTIVE_EXAMS.filter(e=>e.semester===sem);
  const semScores = semExams.map(ex => getScores(studentId, ex.id));
  const semTotals = semScores.map(sc => getTotal(sc));
  const hasAny    = semTotals.some(v=>v!==null);

  // 跨學期總分折線資料（全部）
  const allExams  = ACTIVE_EXAMS;
  const allScores = allExams.map(ex => getScores(studentId, ex.id));
  const allTotals = allScores.map(sc => getTotal(sc));

  // ── 自動文字摘要 ──────────────────────────────────────
  let summaryHtml = "";
  if (hasAny) {
    const validPairs = semTotals.map((t,i)=>({t,sc:semScores[i],ex:semExams[i]})).filter(x=>x.t!==null);
    const first = validPairs[0], last = validPairs[validPairs.length-1];
    const diff  = validPairs.length>=2 ? last.t - first.t : null;

    // 最新有資料的成績
    const latestSc = last.sc;
    const filledSubs = ACTIVE_SUBJECTS.filter(s=>latestSc[s]!==undefined&&latestSc[s]!=="");
    const bestSub  = filledSubs.length ? filledSubs.reduce((a,b)=>parseFloat(latestSc[b])>parseFloat(latestSc[a])?b:a) : null;
    const worstSub = filledSubs.length ? filledSubs.reduce((a,b)=>parseFloat(latestSc[b])<parseFloat(latestSc[a])?b:a) : null;
    const failSubs = filledSubs.filter(s=>parseFloat(latestSc[s])<60);
    const latestRank = latestSc["班排"]||null;
    const firstRank  = first.sc["班排"]||null;
    const rankDiff   = (latestRank&&firstRank&&validPairs.length>=2) ? firstRank-latestRank : null;

    let lines = [];
    if (diff !== null) {
      lines.push(diff>0
        ? `📈 ${sem}學期整體進步 <strong class="up">+${diff.toFixed(0)} 分</strong>`
        : diff<0
          ? `📉 ${sem}學期整體退步 <strong class="down">${diff.toFixed(0)} 分</strong>`
          : `➡️ ${sem}學期成績與期初持平`);
    } else {
      lines.push(`📊 ${sem}學期目前已有 <strong>${validPairs.length}</strong> 次段考資料`);
    }
    if (bestSub)  lines.push(`💪 強項科目：<strong>${bestSub}</strong>（${parseFloat(latestSc[bestSub]).toFixed(0)} 分）`);
    if (worstSub && worstSub!==bestSub) {
      const worstScore = parseFloat(latestSc[worstSub]).toFixed(0);
      lines.push(parseFloat(worstScore)<60
        ? `⚠️ 需加強：<strong class="warn">${worstSub}</strong>（${worstScore} 分，未達及格）`
        : `📌 較弱科目：<strong>${worstSub}</strong>（${worstScore} 分）`);
    }
    if (failSubs.length>0) lines.push(`🔴 不及格科目：<strong class="down">${failSubs.join("、")}</strong>`);
    // 連續兩次低於60的科目警示
    const warnSubs = ACTIVE_SUBJECTS.filter(sub => {
      const vals = validPairs.map(p => p.sc[sub]!==undefined&&p.sc[sub]!==""?parseFloat(p.sc[sub]):null).filter(v=>v!==null);
      return vals.length>=2 && vals[vals.length-1]<60 && vals[vals.length-2]<60;
    });
    if (warnSubs.length>0) lines.push(`🚨 連續不及格：<strong class="down">${warnSubs.join("、")}</strong>（需要特別加強）`);

    // 個人 vs 班級比較
    const lastExForComp = semExams[semExams.length-1];
    const vsClass = ACTIVE_SUBJECTS.map(sub => {
      const myV = parseFloat(latestSc[sub]);
      const clsV = getExamSubjectAvg(lastExForComp.id, sub);
      if (isNaN(myV) || clsV === null) return null;
      return { sub, diff: myV - clsV };
    }).filter(Boolean);
    const aboveAvg = vsClass.filter(x=>x.diff>=5).sort((a,b)=>b.diff-a.diff).slice(0,3);
    const belowAvg = vsClass.filter(x=>x.diff<=-5).sort((a,b)=>a.diff-b.diff).slice(0,3);
    if (aboveAvg.length) lines.push(`📊 高於班平均：<strong>${aboveAvg.map(x=>x.sub+"(+"+x.diff.toFixed(1)+")").join("、")}</strong>`);
    if (belowAvg.length) lines.push(`📊 低於班平均：<strong class="down">${belowAvg.map(x=>x.sub+"("+x.diff.toFixed(1)+")").join("、")}</strong>`);
    if (rankDiff!==null) lines.push(rankDiff>0
      ? `🏅 班排進步 <strong class="up">${rankDiff} 名</strong>（${firstRank} → ${latestRank}）`
      : rankDiff<0
        ? `📉 班排退步 <strong class="down">${Math.abs(rankDiff)} 名</strong>（${firstRank} → ${latestRank}）`
        : `🏅 班排維持第 <strong>${latestRank}</strong> 名`);

    summaryHtml = `
      <div class="summary-box mb-16">
        <div class="summary-title">📋 ${sem}學期學習摘要</div>
        <div class="summary-text">${lines.join("<br>")}</div>
      </div>`;
  }

  const chartsHtml = `
    <div class="grid2 mb-16">
      <div class="card"><div class="card-title">📈 ${sem}學期總分曲線</div><div class="chart-box" style="height:200px"><canvas id="chart-sem-trend"></canvas></div></div>
      <div class="card"><div class="card-title">🕸 科目雷達圖（學期平均）</div><div class="chart-box" style="height:200px"><canvas id="chart-radar"></canvas></div></div>
    </div>
    <div class="card mb-16"><div class="card-title">📊 跨學期總分走勢</div><div class="chart-box" style="height:180px"><canvas id="chart-all-trend"></canvas></div></div>
  `;

  // ── 各次段考卡片 ──────────────────────────────────────
  let examCardsHtml = `<div class="mb-16">`;
  const BADGE_COLORS = [["#EAF1F8","#2D5F8A"],["#EDF4EA","#2E5A1A"],["#FDF0E6","#8B4A14"]];

  semExams.forEach((ex, i) => {
    const sc     = semScores[i];
    const total  = semTotals[i];
    const avgVal = getAvg(sc);
    const filled = getFilledCount(sc);
    const prevSc = i>0 ? semScores[i-1] : null;
    const [bgC, fgC] = BADGE_COLORS[i%3];

    // 各科分數列
    let subjectRows = "";
    if (filled > 0) {
      subjectRows = ACTIVE_SUBJECTS.map(sub => {
        const val  = sc[sub]!==undefined&&sc[sub]!=="" ? parseFloat(sc[sub]) : null;
        const prev = prevSc&&prevSc[sub]!==undefined&&prevSc[sub]!=="" ? parseFloat(prevSc[sub]) : null;
        const diff = (val!==null&&prev!==null) ? val-prev : null;
        // 連續兩次低於60警示
        const prevVal2 = prevSc&&prevSc[sub]!==undefined&&prevSc[sub]!==""?parseFloat(prevSc[sub]):null;
        const isConsecFail = val!==null&&val<60&&prevVal2!==null&&prevVal2<60;
        const barColor = val===null?"#E0DAD0":val>=80?"#5B8A4A":val<60?"#A83232":"#2D5F8A";
        const diffHtml = diff===null?""
          : diff>0?`<span class="subject-diff diff-up">▲${diff.toFixed(0)}</span>`
          : diff<0?`<span class="subject-diff diff-down">▼${Math.abs(diff).toFixed(0)}</span>`
          : `<span class="subject-diff diff-same">—</span>`;
        const subRank = val!==null ? getSubjectRank(studentId, ex.id, sub) : null;
        const subPR   = val!==null ? getSubjectPercentile(studentId, ex.id, sub) : null;
        const subRankHtml = subRank!==null
          ? `<span style="font-size:10px;color:#8B7355;min-width:36px;text-align:right">班${subRank}名</span>`
          : `<span style="min-width:36px"></span>`;
        const prColor = subPR===null?"":subPR>=80?"#2E5A1A":subPR>=60?"#2D5F8A":subPR>=40?"#C4651A":"#8B2222";
        const prHtml  = subPR!==null
          ? `<span title="PR值：高於班上 ${subPR}% 的同學" style="font-size:9px;padding:1px 5px;border-radius:99px;background:${prColor}18;color:${prColor};font-weight:700;min-width:40px;text-align:center;cursor:default">PR${subPR}</span>`
          : `<span style="min-width:40px"></span>`;
        return `<div class="subject-row" style="${isConsecFail?'background:#FFF0F0;border-radius:4px;margin:1px 0;':''}">
          <span class="subject-name">${sub}</span>
          <div class="subject-bar-wrap"><div class="subject-bar-fill" style="width:${val!==null?val:0}%;background:${barColor}"></div></div>
          <span class="subject-score" style="color:${val!==null?(val>=80?"#2E5A1A":val<60?"#8B2222":"#1C1A14"):"#C8BA9E"}">${val!==null?val.toFixed(0):"—"}</span>
          ${subRankHtml}
          ${prHtml}
          ${diffHtml}
        </div>`;
      }).join("");
    }

    const hasData = filled > 0;
    examCardsHtml += `
      <div class="exam-card">
        <div class="exam-card-header" onclick="toggleExamCard('ec-${i}')">
          <div class="exam-card-badge" style="background:${bgC};color:${fgC}">${i+1}</div>
          <div class="exam-card-info">
            <div class="exam-card-name">${ex.name}</div>
            <div class="exam-card-meta">${hasData?`已填 ${filled}/${ACTIVE_SUBJECTS.length} 科`:'尚無資料'}</div>
          </div>
          <div style="text-align:right">
            <div class="exam-card-total">${total!==null?total.toFixed(0):"—"}</div>
            <div class="exam-card-avg">${avgVal!==null?"平均 "+avgVal.toFixed(1):""}</div>
          </div>
          <div class="exam-card-arrow" id="arrow-ec-${i}">▶</div>
        </div>
        <div class="exam-card-body" id="ec-${i}">
          ${hasData ? subjectRows : '<div class="empty-state small"><div class="empty-text">尚未輸入此次段考成績</div></div>'}
          ${sc["班排"]?`<div style="font-size:12px;color:#6B5F4A;margin-top:10px;padding-top:8px;border-top:1px solid #F0EAE0">班排：<strong>${sc["班排"]}</strong> 名　校排：<strong>${sc["校排"]||"—"}</strong> 名</div>`:""}
        </div>
      </div>`;
  });
  examCardsHtml += "</div>";

  // ── 個人科目×段考熱力圖 ───────────────────────────────────
  const allPersonalExams = ACTIVE_EXAMS.filter(ex => getFilledCount(getScores(studentId, ex.id)) > 0);
  let personalHeatmapHtml = "";
  if (allPersonalExams.length >= 2) {
    function personalScoreColor(val) {
      if (val===null) return "#F5F0E8";
      if (val>=90) return "#1A5C2A"; if (val>=80) return "#2E7A3A";
      if (val>=70) return "#5B8A4A"; if (val>=60) return "#C4651A";
      if (val>=50) return "#D4822A"; return "#A83232";
    }
    let hmHtml = `<div style="overflow-x:auto;-webkit-overflow-scrolling:touch">
      <table style="border-collapse:separate;border-spacing:2px;min-width:500px">
      <thead><tr>
        <th style="padding:4px 8px;text-align:left;font-size:11px;color:#6B5F4A;font-weight:600;white-space:nowrap">科目</th>`;
    allPersonalExams.forEach(ex => {
      hmHtml += `<th style="padding:4px 3px;text-align:center;font-size:10px;color:#6B5F4A;font-weight:600;white-space:nowrap">${ex.name.replace("次段考","").replace("第","")}</th>`;
    });
    hmHtml += `<th style="padding:4px 8px;text-align:center;font-size:11px;color:#6B5F4A;font-weight:600">均分</th></tr></thead><tbody>`;

    ACTIVE_SUBJECTS.forEach(sub => {
      const vals = allPersonalExams.map(ex => {
        const v = getScores(studentId, ex.id)[sub];
        return v!==undefined&&v!==""?parseFloat(v):null;
      });
      const filled = vals.filter(v=>v!==null);
      if (!filled.length) return;
      const subAvg = filled.reduce((a,b)=>a+b,0)/filled.length;
      const clsAvgs = allPersonalExams.map(ex => getExamSubjectAvg(ex.id, sub));
      hmHtml += `<tr><td style="padding:3px 8px;font-size:12px;font-weight:600;white-space:nowrap;color:#1C1A14">${sub}</td>`;
      vals.forEach((val, vi) => {
        const clsAvg = clsAvgs[vi];
        const vsAvg = (val!==null&&clsAvg!==null)?val-clsAvg:null;
        const vsHtml = vsAvg!==null
          ? `<div style="font-size:8px;color:${vsAvg>=0?"#A8D8A8":"#F4A0A0"};line-height:1">${vsAvg>=0?"+":""}${vsAvg.toFixed(0)}</div>`
          : "";
        hmHtml += `<td style="padding:2px;text-align:center"><div title="${val!==null?val.toFixed(0):"無資料"}${vsAvg!==null?(vsAvg>=0?" (+"+vsAvg.toFixed(1)+"vs班平均)":" ("+vsAvg.toFixed(1)+"vs班平均)"):""}" style="background:${personalScoreColor(val)};color:#fff;border-radius:4px;padding:4px 2px;font-size:11px;font-family:'DM Mono',monospace;font-weight:600;min-width:28px">${val!==null?val.toFixed(0):"·"}${vsHtml}</div></td>`;
      });
      const avgColor = subAvg>=80?"#2E7A3A":subAvg>=70?"#5B8A4A":subAvg>=60?"#C4651A":"#A83232";
      hmHtml += `<td style="padding:3px 6px;text-align:center;font-family:'DM Mono',monospace;font-weight:700;font-size:12px;color:${avgColor}">${subAvg.toFixed(1)}</td></tr>`;
    });

    // 總分列
    const totalsRow = allPersonalExams.map(ex => getTotal(getScores(studentId, ex.id)));
    hmHtml += `<tr style="border-top:2px solid #C8BA9E;background:#FAF7F0">
      <td style="padding:4px 8px;font-size:11px;font-weight:700;color:#6B5F4A">總分</td>`;
    totalsRow.forEach(t => {
      const tColor = t===null?"#C8BA9E":t>=ACTIVE_SUBJECTS.length*80?"#2E5A1A":t<ACTIVE_SUBJECTS.length*60?"#8B2222":"#2D5F8A";
      hmHtml += `<td style="padding:2px;text-align:center"><div style="background:${tColor};color:#fff;border-radius:4px;padding:4px 2px;font-size:11px;font-family:'DM Mono',monospace;font-weight:700;min-width:28px">${t!==null?t.toFixed(0):"·"}</div></td>`;
    });
    const validT = totalsRow.filter(v=>v!==null);
    const avgT = validT.length?validT.reduce((a,b)=>a+b,0)/validT.length:null;
    hmHtml += `<td style="padding:4px 6px;text-align:center;font-family:'DM Mono',monospace;font-weight:700;font-size:12px;color:#2D5F8A">${avgT!==null?avgT.toFixed(1):"—"}</td></tr>`;
    hmHtml += `</tbody></table></div>
    <div style="font-size:11px;color:#9E9890;margin-top:6px">格內小字為與班平均的差距（紅字=低於班平均，綠字=高於班平均）</div>`;

    personalHeatmapHtml = `
      <div class="card mb-16">
        <div class="card-title">🟥 個人科目熱力圖
          <span style="font-size:11px;font-weight:400;color:#9E9890;margin-left:8px">顏色深淺代表分數高低，含與班平均差距</span>
        </div>
        ${hmHtml}
      </div>`;
  }
  const memoText = S.studentMemos[studentId] || "";
  const memoHtml = `
    <div class="card mb-16" id="student-memo-card">
      <div class="card-title" style="display:flex;justify-content:space-between;align-items:center">
        <span>📝 導師備忘錄</span>
        <span style="font-size:11px;font-weight:400;color:#9E9890">僅導師可見，自動儲存</span>
      </div>
      <div id="student-memo-box"
           contenteditable="true"
           spellcheck="false"
           data-placeholder="點此輸入此學生的私人備忘錄（例如：家庭狀況、已約談記錄、特殊需求...）"
           style="min-height:72px;padding:10px;border:1px solid #C8BA9E;border-radius:6px;font-size:13px;color:#1C1A14;outline:none;cursor:text;line-height:1.7;transition:border .15s"
           onfocus="this.style.borderColor='#8B7355';this.style.background='#FDFBF7'"
           onblur="this.style.borderColor='#C8BA9E';this.style.background=''"
      ></div>
    </div>`;

  wrap.innerHTML = summaryHtml + chartsHtml + examCardsHtml + personalHeatmapHtml + memoHtml;

  // 載入備忘錄內容並綁定自動儲存
  const memoEl = $("student-memo-box");
  if (memoEl) {
    if (memoText) memoEl.innerText = memoText;
    let memoTimer;
    memoEl.oninput = () => {
      clearTimeout(memoTimer);
      memoTimer = setTimeout(() => saveStudentMemo(studentId, memoEl.innerText.trim()), 1500);
    };
  }

  // ── 渲染圖表 ─────────────────────────────────────────
  // 學期折線（含班排第二 y 軸）— 動態 Y 軸
  const validTotals = semTotals.filter(v=>v!==null);
  const yMin = validTotals.length ? Math.max(0, Math.min(...validTotals) - 30) : 0;
  const yMax = validTotals.length ? Math.min(ACTIVE_SUBJECTS.length*100, Math.max(...validTotals) + 30) : ACTIVE_SUBJECTS.length*100;
  const rankData = semExams.map((ex,i) => {
    const r = semScores[i]["班排"];
    return r ? parseInt(r) : null;
  });
  const hasRank = rankData.some(v=>v!==null);
  destroyChart("chart-sem-trend");
  chartInstances["chart-sem-trend"] = new Chart($("chart-sem-trend"), {
    type:"line",
    data:{ labels: semExams.map(e=>e.name),
      datasets:[
        { label:"總分", data:semTotals, yAxisID:"yTotal",
          borderColor:"#2D5F8A", backgroundColor:"#2D5F8A22",
          borderWidth:2.5, pointRadius:6, pointBackgroundColor:"#2D5F8A", fill:true, tension:0.3, spanGaps:true },
        ...(hasRank?[{ label:"班排名", data:rankData, yAxisID:"yRank",
          borderColor:"#C4651A", backgroundColor:"transparent",
          borderWidth:2, pointRadius:5, borderDash:[5,3], tension:0.3, spanGaps:true }]:[]),
        ...(() => {
          const srData = semExams.map((ex,i)=>{ const r=semScores[i]["校排"]; return r?parseInt(r):null; });
          return srData.some(v=>v!==null)?[{ label:"校排名", data:srData, yAxisID:"yRank",
            borderColor:"#6B4FA0", backgroundColor:"transparent",
            borderWidth:2, pointRadius:5, borderDash:[3,4], tension:0.3, spanGaps:true }]:[];
        })()
      ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ display:hasRank, position:"bottom", labels:{font:{size:11},boxWidth:12} } },
      scales:{
        yTotal:{ type:"linear", position:"left", min:yMin, max:yMax, grid:{color:"#E2DED6"}, ticks:{font:{size:11}},
          title:{display:true,text:"總分",font:{size:10},color:"#9E9890"} },
        yRank:{ type:"linear", position:"right", reverse:true,
          ticks:{font:{size:10},stepSize:1}, grid:{display:false},
          title:{display:true,text:"班排名",font:{size:10},color:"#C4651A"} },
        x:{ grid:{display:false}, ticks:{font:{size:11}} }
      }
    }
  });

  // 跨學期折線（動態 Y 軸 + 校排第二軸）
  const allValidTotals = allTotals.filter(v=>v!==null);
  const allYMin = allValidTotals.length ? Math.max(0, Math.min(...allValidTotals) - 30) : 0;
  const allYMax = allValidTotals.length ? Math.min(ACTIVE_SUBJECTS.length*100, Math.max(...allValidTotals) + 30) : ACTIVE_SUBJECTS.length*100;
  const allRankData = allExams.map(ex => {
    const sc = allScores[ACTIVE_EXAMS.indexOf(ex)];
    return sc && sc["校排"] ? parseInt(sc["校排"]) : null;
  });
  const hasAllRank = allRankData.some(v=>v!==null);
  const allRankMax = hasAllRank ? Math.max(...allRankData.filter(v=>v!==null)) + 5 : 50;

  destroyChart("chart-all-trend");
  const allLabels = allExams.map(e=>e.name);
  chartInstances["chart-all-trend"] = new Chart($("chart-all-trend"), {
    type:"line",
    data:{ labels:allLabels,
      datasets:[
        { label:"總分", data:allTotals,
          borderColor:"#C4651A", backgroundColor:"#C4651A15",
          borderWidth:2, pointRadius:4,
          pointBackgroundColor: allExams.map(e=>e.semester===sem?"#C4651A":"#C8BA9E"),
          pointRadius: allExams.map(e=>e.semester===sem?6:3),
          fill:true, tension:0.3, spanGaps:true, yAxisID:"yTotal" },
        ...(hasAllRank ? [{ label:"校排", data:allRankData,
          borderColor:"#6B4FA0", backgroundColor:"transparent",
          borderWidth:1.5, pointRadius:3, borderDash:[5,3],
          tension:0.3, spanGaps:true, yAxisID:"yRank" }] : [])
      ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ display:hasAllRank, position:"bottom", labels:{font:{size:10},boxWidth:10,padding:8} },
        annotation:{} },
      scales:{
        yTotal:{ type:"linear", position:"left", min:allYMin, max:allYMax,
          grid:{color:"#E2DED6"}, ticks:{font:{size:10}},
          title:{display:true,text:"總分",font:{size:10},color:"#C4651A"} },
        ...(hasAllRank ? { yRank:{ type:"linear", position:"right", reverse:true,
          min:1, max:allRankMax,
          grid:{drawOnChartArea:false}, ticks:{font:{size:10},color:"#6B4FA0"},
          title:{display:true,text:"校排",font:{size:10},color:"#6B4FA0"} } } : {}),
        x:{ grid:{display:false}, ticks:{font:{size:9}, maxRotation:45} }
      }
    }
  });

  // 雷達圖（個人學期平均 + 班級平均疊加）
  const semSubAvgs = ACTIVE_SUBJECTS.map(sub => {
    const vals = semScores.map(sc=>sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null).filter(v=>v!==null);
    return { sub, avg: vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null };
  }).filter(s=>s.avg!==null);
  const radarSubs = semSubAvgs.map(s=>s.sub);
  const radarData = semSubAvgs.map(s=>s.avg);

  // 班級平均（各科，以最新有資料的段考為準）
  const latestExForRadar = semExams.slice().reverse().find(ex => getFilledCount(getScores(studentId, ex.id)) > 0);
  const classAvgData = radarSubs.map(sub => {
    const examForAvg = latestExForRadar || semExams[semExams.length-1];
    return getExamSubjectAvg(examForAvg.id, sub);
  });
  const hasClassData = classAvgData.some(v=>v!==null);

  destroyChart("chart-radar");
  chartInstances["chart-radar"] = new Chart($("chart-radar"), {
    type:"radar",
    data:{ labels:radarSubs, datasets:[
      { label: st.name + " " + sem + "平均",
        data: radarData,
        borderColor:"#6B4FA0", backgroundColor:"#6B4FA022", borderWidth:2.5, pointRadius:4 },
      ...(hasClassData ? [{ label:"班級平均",
        data: classAvgData,
        borderColor:"#C4651A", backgroundColor:"#C4651A0D",
        borderWidth:1.5, pointRadius:3, borderDash:[4,3] }] : [])
    ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{legend:{display:hasClassData, position:"bottom", labels:{font:{size:10},boxWidth:10,padding:6}}},
      scales:{ r:{min:0, max:100, ticks:{font:{size:10},stepSize:20}, pointLabels:{font:{size:11}}} }
    }
  });

  // 預設展開第一張有資料的卡片
  // 展開所有有資料的段考卡片
  semScores.forEach((sc, i) => {
    if (getFilledCount(sc) > 0) toggleExamCard("ec-" + i);
  });
}

function toggleExamCard(id) {
  const body  = document.getElementById(id);
  const arrow = document.getElementById("arrow-"+id);
  if (!body) return;
  const isOpen = body.classList.contains("open");
  body.classList.toggle("open", !isOpen);
  if (arrow) arrow.classList.toggle("open", !isOpen);
}


// ── 報告頁 ────────────────────────────────────────────────────
function onReportStudentChange() {
  S.reportStudentId = $("report-student").value;
  const pns = $("parent-notice-student");
  if (pns) pns.value = S.reportStudentId;
  renderReport();
}

function renderReport() {
  // 同步更新家長通知單
  setTimeout(() => { if ($("parent-notice-wrap") && S.reportStudentId) renderParentNotice(); }, 100);
  const studentId = S.reportStudentId;
  const examFilter = $("report-exam").value;
  const hideRank   = $("report-hide-rank").checked;
  const semSummaryMode = $("report-semester-summary")?.checked;
  const wrap = $("report-content");

  if (!studentId) {
    wrap.innerHTML = `<div class="empty-state"><div class="empty-icon">📄</div><div class="empty-title">請選擇學生以產出報告</div></div>`;
    return;
  }

  const st = S.students.find(s=>s.id===studentId); if (!st) return;
  const reportSem = $("report-sem")?.value || "";

  // ── 學期綜合報告模式 ──────────────────────────────────────
  if (semSummaryMode) {
    // 若學期未選，自動選第一個有資料的學期
    let reportSemForSummary = reportSem;
    if (!reportSemForSummary && studentId) {
      reportSemForSummary = SEMESTERS.find(sem =>
        ACTIVE_EXAMS.filter(e=>e.semester===sem).some(ex => getFilledCount(getScores(studentId, ex.id)) > 0)
      ) || "";
      // 同步回 UI
      const semEl = $("report-sem");
      if (semEl && reportSemForSummary) {
        semEl.value = reportSemForSummary;
        // 同步更新段考下拉選單
        buildExamOptions("report-exam", reportSemForSummary);
      }
    }
    if (!reportSemForSummary) {
      wrap.innerHTML = `<div class="empty-state"><div class="empty-icon">📋</div><div class="empty-title">尚無任何成績資料</div><div class="empty-text">請先輸入成績再產出學期綜合報告</div></div>`;
      return;
    }
    renderSemesterSummaryReport(studentId, st, reportSemForSummary, hideRank);
    return;
  }

  if (!studentId) {
    wrap.innerHTML = `<div class="empty-state"><div class="empty-icon">📄</div><div class="empty-title">請選擇學生以產出報告</div></div>`;
    return;
  }
  const examsToShow = examFilter !== "" ? ACTIVE_EXAMS.filter(e=>e.id===examFilter)
    : reportSem ? ACTIVE_EXAMS.filter(e=>e.semester===reportSem) : ACTIVE_EXAMS;
  const allScores   = ACTIVE_EXAMS.map(ex => getScores(studentId, ex.id));

  let html = `
    <div class="report-paper">
      <div class="report-header">
        <div>
          <div class="report-school">${getClassName()} 成績報告 · ${getClassYear()}</div>
          <div class="report-name">${st.name}</div>
          <div class="report-meta">座號：${st.number||"—"} ／ 列印日期：${new Date().toLocaleDateString("zh-TW")}</div>
        </div>
        <div style="font-size:28px">📚</div>
      </div>
  `;

  // 找前一次有資料的段考（用於進退步比較）
  function getPrevScore(exIdx, sub) {
    for (let i = exIdx-1; i >= 0; i--) {
      const v = allScores[i][sub];
      if (v!==undefined && v!=="") return parseFloat(v);
    }
    return null;
  }
  function getPrevTotal(exIdx) {
    for (let i = exIdx-1; i >= 0; i--) {
      const t = getTotal(allScores[i]);
      if (t!==null) return t;
    }
    return null;
  }

  examsToShow.forEach(ex => {
    const exIdx = ACTIVE_EXAMS.findIndex(e=>e.id===ex.id);
    const sc    = allScores[exIdx];
    const total = getTotal(sc);
    const prevTotal = getPrevTotal(exIdx);
    const totalDiff = (total!==null && prevTotal!==null) ? total - prevTotal : null;

    // 各科班平均（用於與班平均比較）
    const subClassAvgs = {};
    ACTIVE_SUBJECTS.forEach(sub => { subClassAvgs[sub] = getExamSubjectAvg(ex.id, sub); });
    const classAvgTotal = (() => {
      const ts = S.students.map(st2=>getTotal(getScores(st2.id,ex.id))).filter(v=>v!==null);
      return ts.length ? ts.reduce((a,b)=>a+b,0)/ts.length : null;
    })();

    html += `
      <div class="report-section">
        <div class="report-exam-title" style="display:flex;justify-content:space-between;align-items:center">
          <span>${ex.name}</span>
          ${!hideRank&&sc["班排"]?`<span style="font-size:12px;color:#6B5F4A;font-weight:400">班排 <strong>${sc["班排"]}</strong> 名　校排 <strong>${sc["校排"]||"—"}</strong> 名</span>`:""}
        </div>
        <table class="report-table">
          <thead><tr>
            <th style="width:56px">科目</th>
            <th style="width:52px;text-align:center">分數</th>
            <th style="min-width:100px">表現</th>
            <th style="width:52px;text-align:center">進退步</th>
            ${!hideRank?'<th style="width:56px;text-align:center">班排名</th>':""}
            <th style="width:64px;text-align:center">班平均</th>
            <th style="width:52px;text-align:center">差距</th>
          </tr></thead>
          <tbody>
    `;
    ACTIVE_SUBJECTS.forEach(sub => {
      const v      = sc[sub];
      const hasVal = v!==undefined && v!=="";
      const pct    = hasVal ? parseFloat(v) : null;
      const prev   = getPrevScore(exIdx, sub);
      const diff   = (pct!==null && prev!==null) ? pct - prev : null;
      const clsAvg = subClassAvgs[sub];
      const vsAvg  = (pct!==null && clsAvg!==null) ? pct - clsAvg : null;

      // 級距標籤
      let lvlLabel="", lvlBg="", lvlColor="";
      if (pct!==null) {
        if(pct>=90){lvlLabel="優秀";lvlBg="#EDF4EA";lvlColor="#2E5A1A";}
        else if(pct>=80){lvlLabel="良好";lvlBg="#E8F4F0";lvlColor="#1A6B4A";}
        else if(pct>=70){lvlLabel="中等";lvlBg="#EAF1F8";lvlColor="#1C4A6B";}
        else if(pct>=60){lvlLabel="及格";lvlBg="#FDF0E6";lvlColor="#8B4A14";}
        else{lvlLabel="待加強";lvlBg="#FAECEC";lvlColor="#8B2222";}
      }

      const barColor = pct!==null?(pct>=80?"#5B8A4A":pct<60?"#A83232":"#2D5F8A"):"#E0DAD0";
      const subRank  = (!hideRank && pct!==null) ? getSubjectRank(studentId, ex.id, sub) : null;
      const isConsecFail = pct!==null && pct<60 && prev!==null && prev<60;

      const diffHtml = diff===null ? `<span style="color:#C8BA9E">—</span>`
        : diff>0 ? `<span style="color:#2E5A1A;font-weight:600">▲${diff.toFixed(0)}</span>`
        : diff<0 ? `<span style="color:#8B2222;font-weight:600">▼${Math.abs(diff).toFixed(0)}</span>`
        : `<span style="color:#9E9890">持平</span>`;

      const vsAvgHtml = vsAvg===null ? `<span style="color:#C8BA9E">—</span>`
        : vsAvg>=0 ? `<span style="color:#2E5A1A;font-weight:600">+${vsAvg.toFixed(1)}</span>`
        : `<span style="color:#8B2222;font-weight:600">${vsAvg.toFixed(1)}</span>`;

      html += `<tr style="${isConsecFail?'background:#FFF8F8':''}">
        <td style="font-weight:${isConsecFail?'700':'400'};color:${isConsecFail?'#8B2222':'inherit'}">${sub}${isConsecFail?'⚠️':''}</td>
        <td style="text-align:center;font-family:monospace;font-weight:700;font-size:14px;color:${pct!==null?(pct>=80?"#2E5A1A":pct<60?"#8B2222":"#1C1A14"):"#9E9890"}">${hasVal?pct.toFixed(0):"—"}</td>
        <td>
          ${pct!==null?`
          <div style="display:flex;align-items:center;gap:6px">
            <div style="flex:1;height:7px;background:#F0EEE9;border-radius:3px;overflow:hidden">
              <div style="height:100%;width:${pct}%;background:${barColor};border-radius:3px"></div>
            </div>
            <span style="font-size:10px;padding:1px 6px;border-radius:99px;background:${lvlBg};color:${lvlColor};font-weight:600;white-space:nowrap">${lvlLabel}</span>
          </div>`:"—"}
        </td>
        <td style="text-align:center;font-size:12px">${diffHtml}</td>
        ${!hideRank?`<td style="text-align:center;font-size:11px;color:#6B5F4A">${subRank?`第${subRank}名`:"—"}</td>`:""}
        <td style="text-align:center;font-size:11px;font-family:monospace;color:#6B5F4A">${clsAvg!==null?clsAvg.toFixed(1):"—"}</td>
        <td style="text-align:center;font-size:12px">${vsAvgHtml}</td>
      </tr>`;
    });

    // 總分列
    const totalDiffHtml = totalDiff===null ? ""
      : totalDiff>0 ? `<span style="color:#2E5A1A;font-size:11px;margin-left:6px">▲${totalDiff.toFixed(0)}</span>`
      : totalDiff<0 ? `<span style="color:#8B2222;font-size:11px;margin-left:6px">▼${Math.abs(totalDiff).toFixed(0)}</span>`
      : "";
    const vsTotalAvg = (total!==null&&classAvgTotal!==null) ? total-classAvgTotal : null;

    html += `</tbody>
          <tfoot><tr style="background:#FAF7F0">
            <td style="font-weight:700">總分</td>
            <td style="text-align:center;font-weight:700;font-family:monospace;font-size:15px">${total!==null?total.toFixed(0):"—"}${totalDiffHtml}</td>
            <td colspan="${hideRank?2:3}" style="font-size:11px;color:#6B5F4A">
              ${getAvg(sc)!==null?`平均分 ${getAvg(sc).toFixed(1)} 分　填寫科數 ${getFilledCount(sc)}/${ACTIVE_SUBJECTS.length} 科`:""}
            </td>
            <td style="text-align:center;font-size:11px;font-family:monospace;color:#6B5F4A">${classAvgTotal!==null?classAvgTotal.toFixed(1):"—"}</td>
            <td style="text-align:center;font-size:12px">${vsTotalAvg===null?"—":vsTotalAvg>=0?`<span style="color:#2E5A1A;font-weight:600">+${vsTotalAvg.toFixed(1)}</span>`:`<span style="color:#8B2222;font-weight:600">${vsTotalAvg.toFixed(1)}</span>`}</td>
          </tr></tfoot>
        </table>
      </div>
    `;
  });

  // ── 評語：從 S.teacherComments 讀取（跨學生、跨切換保存）
  const examIdForComment = examFilter || "all";
  const savedComment = (S.teacherComments[studentId] || {})[examIdForComment] || "";

  html += `
      <div class="report-comment">
        <div class="report-comment-title">學習摘要（系統自動產生）</div>
        <div id="report-auto-summary" style="font-size:12px;color:#3A3629;line-height:2;padding:8px 0;border-bottom:1px solid #E0DAD0;margin-bottom:10px"></div>
        <div class="report-comment-title" style="margin-top:8px">導師評語
          <span style="font-size:10px;font-weight:400;color:#9E9890;margin-left:6px">（可直接點擊輸入，自動儲存）</span>
        </div>
        <div id="report-teacher-comment"
             class="report-comment-box"
             contenteditable="true"
             spellcheck="false"
             style="color:#1C1A14;outline:none;cursor:text"
             data-placeholder="點此輸入導師評語..."
        ></div>
      </div>
      <div class="report-sign">
        <div>導師簽名：_______________</div>
        <div>家長簽名：_______________</div>
        <div>日期：_______________</div>
      </div>
    </div>
  `;
  wrap.innerHTML = html;

  // 還原評語
  const commentEl = $("report-teacher-comment");
  if (commentEl && savedComment) commentEl.innerText = savedComment;

  // oninput：存入 S.teacherComments + 同步到家長通知單 + debounce 存 Firebase
  if (commentEl) {
    let saveTimer;
    commentEl.oninput = () => {
      const txt = commentEl.innerText.trim();
      if (!S.teacherComments[studentId]) S.teacherComments[studentId] = {};
      S.teacherComments[studentId][examIdForComment] = txt;
      // 同步到家長通知單
      const pnc = $("parent-notice-teacher-comment");
      if (pnc) pnc.innerText = commentEl.innerText;
      // debounce 2 秒後存 Firebase
      clearTimeout(saveTimer);
      saveTimer = setTimeout(() => saveTeacherComments(), 2000);
    };
  }

  // ── 自動學習摘要 ──────────────────────────────────────
  const summEl = $("report-auto-summary"); if (!summEl) return;
  const rAllScores = ACTIVE_EXAMS.map(ex => getScores(studentId, ex.id));
  const rAllTotals = rAllScores.map(sc => getTotal(sc));
  const rValid = ACTIVE_EXAMS.map((ex,i)=>({ex,sc:rAllScores[i],t:rAllTotals[i]})).filter(x=>x.t!==null);
  if (rValid.length === 0) { summEl.textContent = "尚無成績資料"; return; }
  const rFirst = rValid[0], rLast = rValid[rValid.length-1];
  const rDiff  = rValid.length>=2 ? rLast.t - rFirst.t : null;
  const rLatestSc = rLast.sc;
  const rFilled = ACTIVE_SUBJECTS.filter(s=>rLatestSc[s]!==undefined&&rLatestSc[s]!=="");
  const rBest  = rFilled.length ? rFilled.reduce((a,b)=>parseFloat(rLatestSc[b])>parseFloat(rLatestSc[a])?b:a) : null;
  const rWorst = rFilled.length ? rFilled.reduce((a,b)=>parseFloat(rLatestSc[b])<parseFloat(rLatestSc[a])?b:a) : null;
  const rFail  = rFilled.filter(s=>parseFloat(rLatestSc[s])<60);
  const rWarn  = ACTIVE_SUBJECTS.filter(sub => {
    const vals = rValid.slice(-2).map(p=>p.sc[sub]!==undefined&&p.sc[sub]!==""?parseFloat(p.sc[sub]):null).filter(v=>v!==null);
    return vals.length>=2 && vals.every(v=>v<60);
  });
  let lines2 = [];
  lines2.push(`共累積 ${rValid.length} 次段考資料（${rFirst.ex.name} ～ ${rLast.ex.name}）`);
  if (rDiff!==null) lines2.push(rDiff>0?`整體總分進步 +${rDiff.toFixed(0)} 分`:rDiff<0?`整體總分退步 ${rDiff.toFixed(0)} 分`:"整體成績持平");
  if (rBest)  lines2.push(`最新強項科目：${rBest}（${parseFloat(rLatestSc[rBest]).toFixed(0)} 分）`);
  if (rWorst && rWorst!==rBest) lines2.push(`最新較弱科目：${rWorst}（${parseFloat(rLatestSc[rWorst]).toFixed(0)} 分）`);
  if (rFail.length>0) lines2.push(`不及格科目：${rFail.join("、")}`);
  if (rWarn.length>0) lines2.push(`⚠️ 連續不及格需加強：${rWarn.join("、")}`);
  summEl.innerHTML = lines2.map(l=>`• ${l}`).join("<br>");
}

// ── 全班批次列印報告 ─────────────────────────────────────
function batchPrintAll() {
  if (!S.students.length) { showToast("尚無學生資料"); return; }
  const hideRank = $("report-hide-rank")?.checked || false;
  const semSummaryMode = $("report-semester-summary")?.checked || false;
  const wrap = $("report-content");

  // ── 學期綜合報告批次列印 ──────────────────────────────────
  if (semSummaryMode) {
    const sem = $("report-sem")?.value;
    if (!sem) { showToast("請先選擇學期"); return; }
    const validStudents = S.students.filter(st =>
      ACTIVE_EXAMS.filter(e=>e.semester===sem).some(ex => getFilledCount(getScores(st.id, ex.id)) > 0)
    );
    if (!validStudents.length) { showToast("所選學期沒有任何學生有成績資料"); return; }

    // 逐一渲染學期綜合報告，擷取 HTML 後組合列印
    const origStudentId = S.reportStudentId;
    let idx = 0;
    const pages = [];

    // 直接用 buildSemesterSummaryHTML 產生 HTML 字串，不經過 DOM
    validStudents.forEach((st, i) => {
      const html = buildSemesterSummaryHTML(st.id, st, sem, hideRank);
      if (html) pages.push(html);
    });

    if (!pages.length) { showToast("沒有可列印的資料"); return; }
    wrap.innerHTML = pages.map((html, i) =>
      `<div style="${i>0?'page-break-before:always;':''}">${html}</div>`
    ).join("");
    const noticeSection = document.getElementById("parent-notice-section");
    if (noticeSection) noticeSection.style.display = "none";
    showToast(`⏳ 準備列印 ${pages.length} 位學生的學期綜合報告...`);
    setTimeout(() => {
      window.print();
      setTimeout(() => {
        if (noticeSection) noticeSection.style.display = "";
        S.reportStudentId = origStudentId;
        if (origStudentId) renderReport();
      }, 1000);
    }, 300);
    return;
  }

  // ── 一般批次列印（原有邏輯）──────────────────────────────

  let allHtml = "";
  S.students.forEach((st, idx) => {
    const allScores = ACTIVE_EXAMS.map(ex => getScores(st.id, ex.id));
    // 只顯示有資料的段考
    const examsWithData = ACTIVE_EXAMS.filter((ex,i) => getFilledCount(allScores[i]) > 0);
    if (!examsWithData.length) return;

    allHtml += `<div class="report-paper" style="${idx>0?'page-break-before:always;':''}">
      <div class="report-header">
        <div>
          <div class="report-school">${getClassName()} 成績報告 · ${getClassYear()}</div>
          <div class="report-name">${st.name}</div>
          <div class="report-meta">座號：${st.number||"—"} ／ 列印日期：${new Date().toLocaleDateString("zh-TW")}</div>
        </div>
        <div style="font-size:28px">📚</div>
      </div>`;

    // 批次版的前次總分輔助
    function bGetPrevScore(bAllScores, exIdx, sub) {
      for (let i=exIdx-1;i>=0;i--) { const v=bAllScores[i][sub]; if(v!==undefined&&v!=="") return parseFloat(v); }
      return null;
    }
    function bGetPrevTotal(bAllScores, exIdx) {
      for (let i=exIdx-1;i>=0;i--) { const t=getTotal(bAllScores[i]); if(t!==null) return t; }
      return null;
    }

    examsWithData.forEach(ex => {
      const exIdx = ACTIVE_EXAMS.findIndex(e=>e.id===ex.id);
      const sc    = allScores[exIdx];
      const total = getTotal(sc);
      const prevTotal = bGetPrevTotal(allScores, exIdx);
      const totalDiff = (total!==null&&prevTotal!==null)?total-prevTotal:null;
      const classAvgTotal = (() => { const ts=S.students.map(st2=>getTotal(getScores(st2.id,ex.id))).filter(v=>v!==null); return ts.length?ts.reduce((a,b)=>a+b,0)/ts.length:null; })();

      allHtml += `<div class="report-section">
        <div class="report-exam-title" style="display:flex;justify-content:space-between;align-items:center">
          <span>${ex.name}</span>
          ${!hideRank&&sc["班排"]?`<span style="font-size:11px;color:#6B5F4A;font-weight:400">班排 <strong>${sc["班排"]}</strong> 名　校排 <strong>${sc["校排"]||"—"}</strong> 名</span>`:""}
        </div>
        <table class="report-table"><thead><tr>
          <th style="width:52px">科目</th><th style="width:48px;text-align:center">分數</th>
          <th style="min-width:90px">表現</th><th style="width:48px;text-align:center">進退步</th>
          ${!hideRank?'<th style="width:52px;text-align:center">班排名</th>':""}
          <th style="width:56px;text-align:center">班平均</th><th style="width:48px;text-align:center">差距</th>
        </tr></thead><tbody>`;
      ACTIVE_SUBJECTS.forEach(sub => {
        const v=sc[sub]; const hasVal=v!==undefined&&v!==""; const pct=hasVal?parseFloat(v):null;
        const prev=bGetPrevScore(allScores,exIdx,sub);
        const diff=(pct!==null&&prev!==null)?pct-prev:null;
        const clsAvg=getExamSubjectAvg(ex.id,sub);
        const vsAvg=(pct!==null&&clsAvg!==null)?pct-clsAvg:null;
        const subRank=!hideRank&&pct!==null?getSubjectRank(st.id,ex.id,sub):null;
        const isConsecFail=pct!==null&&pct<60&&prev!==null&&prev<60;
        let lvlLabel="",lvlBg="",lvlColor="";
        if(pct!==null){
          if(pct>=90){lvlLabel="優秀";lvlBg="#EDF4EA";lvlColor="#2E5A1A";}
          else if(pct>=80){lvlLabel="良好";lvlBg="#E8F4F0";lvlColor="#1A6B4A";}
          else if(pct>=70){lvlLabel="中等";lvlBg="#EAF1F8";lvlColor="#1C4A6B";}
          else if(pct>=60){lvlLabel="及格";lvlBg="#FDF0E6";lvlColor="#8B4A14";}
          else{lvlLabel="待加強";lvlBg="#FAECEC";lvlColor="#8B2222";}
        }
        const barColor=pct!==null?(pct>=80?"#5B8A4A":pct<60?"#A83232":"#2D5F8A"):"#E0DAD0";
        const diffHtml=diff===null?`<span style="color:#C8BA9E">—</span>`:diff>0?`<span style="color:#2E5A1A;font-weight:600">▲${diff.toFixed(0)}</span>`:diff<0?`<span style="color:#8B2222;font-weight:600">▼${Math.abs(diff).toFixed(0)}</span>`:`<span style="color:#9E9890">持平</span>`;
        const vsAvgHtml=vsAvg===null?`<span style="color:#C8BA9E">—</span>`:vsAvg>=0?`<span style="color:#2E5A1A;font-weight:600">+${vsAvg.toFixed(1)}</span>`:`<span style="color:#8B2222;font-weight:600">${vsAvg.toFixed(1)}</span>`;
        allHtml += `<tr style="${isConsecFail?'background:#FFF8F8':''}">
          <td style="color:${isConsecFail?'#8B2222':'inherit'};font-weight:${isConsecFail?700:400}">${sub}${isConsecFail?'⚠️':''}</td>
          <td style="text-align:center;font-family:monospace;font-weight:700;font-size:13px;color:${pct!==null?(pct>=80?"#2E5A1A":pct<60?"#8B2222":"#1C1A14"):"#9E9890"}">${hasVal?pct.toFixed(0):"—"}</td>
          <td>${pct!==null?`<div style="display:flex;align-items:center;gap:4px"><div style="flex:1;height:6px;background:#F0EEE9;border-radius:3px;overflow:hidden"><div style="height:100%;width:${pct}%;background:${barColor};border-radius:3px"></div></div><span style="font-size:9px;padding:1px 5px;border-radius:99px;background:${lvlBg};color:${lvlColor};font-weight:600;white-space:nowrap">${lvlLabel}</span></div>`:"—"}</td>
          <td style="text-align:center;font-size:11px">${diffHtml}</td>
          ${!hideRank?`<td style="text-align:center;font-size:10px;color:#6B5F4A">${subRank?`第${subRank}名`:"—"}</td>`:""}
          <td style="text-align:center;font-size:10px;font-family:monospace;color:#6B5F4A">${clsAvg!==null?clsAvg.toFixed(1):"—"}</td>
          <td style="text-align:center;font-size:11px">${vsAvgHtml}</td>
        </tr>`;
      });
      const tdHtml=totalDiff===null?"":totalDiff>0?`<span style="color:#2E5A1A;font-size:10px"> ▲${totalDiff.toFixed(0)}</span>`:`<span style="color:#8B2222;font-size:10px"> ▼${Math.abs(totalDiff).toFixed(0)}</span>`;
      const vsTotalAvg=(total!==null&&classAvgTotal!==null)?total-classAvgTotal:null;
      allHtml += `</tbody><tfoot><tr style="background:#FAF7F0">
        <td style="font-weight:700">總分</td>
        <td style="text-align:center;font-weight:700;font-family:monospace;font-size:14px">${total!==null?total.toFixed(0):"—"}${tdHtml}</td>
        <td colspan="${hideRank?2:3}" style="font-size:10px;color:#6B5F4A">${getAvg(sc)!==null?`平均 ${getAvg(sc).toFixed(1)} 分　${getFilledCount(sc)}/${ACTIVE_SUBJECTS.length} 科`:""}</td>
        <td style="text-align:center;font-size:10px;font-family:monospace;color:#6B5F4A">${classAvgTotal!==null?classAvgTotal.toFixed(1):"—"}</td>
        <td style="text-align:center;font-size:11px">${vsTotalAvg===null?"—":vsTotalAvg>=0?`<span style="color:#2E5A1A;font-weight:600">+${vsTotalAvg.toFixed(1)}</span>`:`<span style="color:#8B2222;font-weight:600">${vsTotalAvg.toFixed(1)}</span>`}</td>
      </tr></tfoot></table></div>`;
    });

    // 自動摘要
    const rValid = ACTIVE_EXAMS.map((ex,i)=>({ex,sc:allScores[i],t:getTotal(allScores[i])})).filter(x=>x.t!==null);
    if (rValid.length) {
      const rFirst=rValid[0], rLast=rValid[rValid.length-1];
      const rDiff = rValid.length>=2 ? rLast.t-rFirst.t : null;
      const rFilled = ACTIVE_SUBJECTS.filter(s=>rLast.sc[s]!==undefined&&rLast.sc[s]!=="");
      const rBest = rFilled.length?rFilled.reduce((a,b)=>parseFloat(rLast.sc[b])>parseFloat(rLast.sc[a])?b:a):null;
      const rFail = rFilled.filter(s=>parseFloat(rLast.sc[s])<60);
      const rWarn = ACTIVE_SUBJECTS.filter(sub=>{
        const vals=rValid.slice(-2).map(p=>p.sc[sub]!==undefined&&p.sc[sub]!==""?parseFloat(p.sc[sub]):null).filter(v=>v!==null);
        return vals.length>=2&&vals.every(v=>v<60);
      });
      let lines=[];
      lines.push(`共 ${rValid.length} 次段考（${rFirst.ex.name} ～ ${rLast.ex.name}）`);
      if(rDiff!==null) lines.push(rDiff>0?`整體進步 +${rDiff.toFixed(0)} 分`:rDiff<0?`整體退步 ${rDiff.toFixed(0)} 分`:"整體持平");
      if(rBest) lines.push(`強項：${rBest}（${parseFloat(rLast.sc[rBest]).toFixed(0)} 分）`);
      if(rFail.length>0) lines.push(`不及格：${rFail.join("、")}`);
      if(rWarn.length>0) lines.push(`⚠️ 連續不及格：${rWarn.join("、")}`);
      allHtml += `<div class="report-comment">
        <div class="report-comment-title">學習摘要</div>
        <div style="font-size:11px;color:#3A3629;line-height:1.8;padding:6px 0;border-bottom:1px solid #E0DAD0;margin-bottom:8px">${lines.map(l=>"• "+l).join("<br>")}</div>
        <div class="report-comment-title" style="margin-top:6px">導師評語</div>
        <div class="report-comment-box">&nbsp;</div>
      </div>
      <div class="report-sign">
        <div>導師簽名：_______________</div>
        <div>家長簽名：_______________</div>
        <div>日期：_______________</div>
      </div>`;
    }
    allHtml += `</div>`;
  });

  if (!allHtml) { showToast("目前沒有任何有成績資料的學生"); return; }
  wrap.innerHTML = allHtml;
  const noticeSec = document.getElementById("parent-notice-section");
  if (noticeSec) noticeSec.style.display = "none";
  setTimeout(() => {
    window.print();
    setTimeout(() => { if (noticeSec) noticeSec.style.display = ""; }, 1000);
  }, 300);
}

function printReport() {
  const wrap = $("report-content");
  if (!wrap || !wrap.querySelector(".report-paper")) { showToast("請先選擇學生"); return; }
  const noticeSection = document.getElementById("parent-notice-section");
  if (noticeSection) noticeSection.style.display = "none";
  window.print();
  setTimeout(() => {
    if (noticeSection) noticeSection.style.display = "";
  }, 800);
}

// ── 列印用 SVG 輔助函式 ───────────────────────────────────────
function makePrintTrendSVG(pairs) {
  if (!pairs || pairs.length < 2) return "";
  const W = 480, H = 80, padL = 30, padR = 10, padT = 10, padB = 20;
  const iW = W - padL - padR, iH = H - padT - padB;
  const vals = pairs.map(p => p.t);
  const minV = Math.max(0, Math.min(...vals) - 20);
  const maxV = Math.min(900, Math.max(...vals) + 20);
  const scaleX = i => padL + (i / (pairs.length - 1 || 1)) * iW;
  const scaleY = v => padT + iH - ((v - minV) / (maxV - minV || 1)) * iH;
  const pts = pairs.map((p, i) => [scaleX(i), scaleY(p.t)]);
  const line = pts.map((p,i) => (i===0?"M":"L") + p[0].toFixed(1) + "," + p[1].toFixed(1)).join(" ");
  const area = line + " L" + pts[pts.length-1][0].toFixed(1) + "," + (padT+iH) + " L" + pts[0][0].toFixed(1) + "," + (padT+iH) + " Z";
  const trend = vals[vals.length-1] - vals[0];
  const color = trend >= 0 ? "#2D5F8A" : "#A83232";
  const labels = pairs.map((p,i) => {
    const x = scaleX(i);
    const lbl = p.ex.name.replace(/第|次段考/g,"").replace("上","↑").replace("下","↓");
    return `<text x="${x.toFixed(1)}" y="${H-4}" text-anchor="middle" font-size="7" fill="#9E9890">${lbl}</text>`;
  }).join("");
  const scoreLabels = pts.map((pt,i) => `<text x="${pt[0].toFixed(1)}" y="${(pt[1]-4).toFixed(1)}" text-anchor="middle" font-size="8" fill="${color}" font-weight="700">${vals[i].toFixed(0)}</text>`).join("");
  const dots = pts.map(pt => `<circle cx="${pt[0].toFixed(1)}" cy="${pt[1].toFixed(1)}" r="3" fill="${color}"/>`).join("");
  return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="overflow:visible">
    <path d="${area}" fill="${color}" fill-opacity="0.08"/>
    <path d="${line}" fill="none" stroke="${color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
    ${dots}${scoreLabels}${labels}
  </svg>`;
}

function makePrintSubjectSVG(sc) {
  const subs = ACTIVE_SUBJECTS.filter(s => sc[s]!==undefined && sc[s]!=="");
  if (!subs.length) return "";
  const W = 480, barH = 14, gap = 4, labelW = 44, scoreW = 28, padR = 8;
  const barW = W - labelW - scoreW - padR;
  const H = subs.length * (barH + gap) + 16;
  const bars = subs.map((sub, i) => {
    const val = parseFloat(sc[sub]);
    const clsAvg = getExamSubjectAvg(
      Object.keys(S.scores).find(k => {
        const sc2 = S.scores[k];
        return sc2[sub] !== undefined && parseFloat(sc2[sub]) === val;
      })?.split("_")[1] || "", sub
    );
    const fillColor = val >= 80 ? "#5B8A4A" : val < 60 ? "#A83232" : "#2D5F8A";
    const y = i * (barH + gap);
    const w = Math.max(0, Math.min(barW, (val/100)*barW));
    return `<text x="${labelW-4}" y="${y+barH-3}" text-anchor="end" font-size="9" fill="#3A3629" font-weight="500">${sub}</text>
      <rect x="${labelW}" y="${y}" width="${barW}" height="${barH}" fill="#F0EAE0" rx="2"/>
      <rect x="${labelW}" y="${y}" width="${w}" height="${barH}" fill="${fillColor}" rx="2"/>
      <text x="${labelW+barW+4}" y="${y+barH-3}" font-size="9" fill="${fillColor}" font-weight="700">${val.toFixed(0)}</text>`;
  }).join("");
  return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg">${bars}</svg>`;
}

// ── 個人分析列印 ──────────────────────────────────────────────
function buildAnalysisPrintPage(studentId, sem) {
  const st = S.students.find(s => s.id === studentId);
  if (!st) return "";
  const date = new Date().toLocaleDateString("zh-TW");

  // 決定要顯示的段考範圍
  let examsToShow, semLabel;
  if (GRADE_MAP[sem]) {
    const gInfo = GRADE_MAP[sem];
    examsToShow = ACTIVE_EXAMS.filter(e => gInfo.semesters.includes(e.semester));
    semLabel = gInfo.label;
  } else {
    examsToShow = ACTIVE_EXAMS.filter(e => e.semester === sem);
    semLabel = sem + "學期";
  }

  const scoresArr = examsToShow.map(ex => getScores(studentId, ex.id));
  const totalsArr = scoresArr.map(sc => getTotal(sc));
  const validPairs = examsToShow.map((ex,i) => ({ex, sc:scoresArr[i], t:totalsArr[i]})).filter(x => x.t !== null);

  if (!validPairs.length) return ""; // 無資料不列印

  const lastSc = validPairs[validPairs.length-1].sc;
  const lastTotal = validPairs[validPairs.length-1].t;
  const firstTotal = validPairs[0].t;
  const diff = validPairs.length >= 2 ? lastTotal - firstTotal : null;
  const filled = ACTIVE_SUBJECTS.filter(s => lastSc[s] !== undefined && lastSc[s] !== "");
  const bestSub = filled.length ? filled.reduce((a,b) => parseFloat(lastSc[b]) > parseFloat(lastSc[a]) ? b : a) : null;
  const failSubs = filled.filter(s => parseFloat(lastSc[s]) < 60);
  const warnSubs = ACTIVE_SUBJECTS.filter(sub => {
    const vals = validPairs.slice(-2).map(p => p.sc[sub]!==undefined&&p.sc[sub]!==""?parseFloat(p.sc[sub]):null).filter(v=>v!==null);
    return vals.length >= 2 && vals.every(v => v < 60);
  });

  // 各科平均（跨所選段考）
  const subAvgs = ACTIVE_SUBJECTS.map(sub => {
    const vals = scoresArr.map(sc => sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null).filter(v=>v!==null);
    return { sub, avg: vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null, count: vals.length };
  }).filter(s => s.avg !== null);

  // 摘要文字
  let lines = [];
  if (diff !== null) lines.push(diff > 0 ? `📈 整體進步 +${diff.toFixed(0)} 分` : diff < 0 ? `📉 整體退步 ${diff.toFixed(0)} 分` : "➡️ 成績持平");
  lines.push(`共 ${validPairs.length} 次有效段考資料`);
  if (bestSub) lines.push(`💪 強項科目：${bestSub}（${parseFloat(lastSc[bestSub]).toFixed(0)} 分）`);
  if (failSubs.length) lines.push(`⚠️ 不及格：${failSubs.join("、")}`);
  if (warnSubs.length) lines.push(`🚨 連續不及格：${warnSubs.join("、")}`);

  // 各次段考成績表
  let examRows = "";
  validPairs.forEach(({ex, sc, t}, vi) => {
    const prevSc = vi > 0 ? validPairs[vi-1].sc : null;
    const prevT  = vi > 0 ? validPairs[vi-1].t  : null;
    const tDiff  = (t !== null && prevT !== null) ? t - prevT : null;
    const tDiffHtml = tDiff === null ? "" : tDiff > 0
      ? `<span style="color:#2E5A1A;font-size:10px"> ▲${tDiff.toFixed(0)}</span>`
      : `<span style="color:#8B2222;font-size:10px"> ▼${Math.abs(tDiff).toFixed(0)}</span>`;

    let subRows = "";
    ACTIVE_SUBJECTS.forEach(sub => {
      const val = sc[sub] !== undefined && sc[sub] !== "" ? parseFloat(sc[sub]) : null;
      const prev = prevSc ? (prevSc[sub]!==undefined&&prevSc[sub]!==""?parseFloat(prevSc[sub]):null) : null;
      const d = (val !== null && prev !== null) ? val - prev : null;
      const clsAvg = getExamSubjectAvg(ex.id, sub);
      const vsAvg  = (val !== null && clsAvg !== null) ? val - clsAvg : null;
      const scoreColor = val === null ? "#9E9890" : val >= 80 ? "#2E5A1A" : val < 60 ? "#8B2222" : "#1C1A14";
      const lvl = val === null ? "" : val >= 90 ? "優秀" : val >= 80 ? "良好" : val >= 70 ? "中等" : val >= 60 ? "及格" : "待加強";
      const dHtml = d === null ? `<span style="color:#C8BA9E">—</span>` : d > 0
        ? `<span style="color:#2E5A1A">▲${d.toFixed(0)}</span>`
        : d < 0 ? `<span style="color:#8B2222">▼${Math.abs(d).toFixed(0)}</span>`
        : `<span style="color:#9E9890">持平</span>`;
      const vsHtml = vsAvg === null ? `<span style="color:#C8BA9E">—</span>`
        : vsAvg >= 0 ? `<span style="color:#2E5A1A">+${vsAvg.toFixed(1)}</span>`
        : `<span style="color:#8B2222">${vsAvg.toFixed(1)}</span>`;
      subRows += `<tr>
        <td style="font-size:11px;padding:3px 6px">${sub}</td>
        <td style="text-align:center;font-family:monospace;font-weight:700;font-size:12px;color:${scoreColor}">${val !== null ? val.toFixed(0) : "—"}</td>
        <td style="text-align:center;font-size:10px">${lvl}</td>
        <td style="text-align:center;font-size:10px">${dHtml}</td>
        <td style="text-align:center;font-size:10px;color:#6B5F4A">${clsAvg !== null ? clsAvg.toFixed(1) : "—"}</td>
        <td style="text-align:center;font-size:10px">${vsHtml}</td>
      </tr>`;
    });

    examRows += `
      <div style="margin-bottom:10px;page-break-inside:avoid">
        <div style="display:flex;justify-content:space-between;align-items:center;background:#1C1A14;color:#F5F0E8;padding:5px 10px;border-radius:4px;font-size:12px;font-weight:700;margin-bottom:4px">
          <span>${ex.name}</span>
          <span style="font-size:11px;font-weight:400">總分 <strong style="font-size:14px">${t !== null ? t.toFixed(0) : "—"}</strong>${tDiffHtml}　班排 ${sc["班排"]||"—"} 名　校排 ${sc["校排"]||"—"} 名</span>
        </div>
        <table style="width:100%;border-collapse:collapse;font-size:11px">
          <thead><tr style="background:#FAF7F0">
            <th style="padding:3px 6px;text-align:left;font-size:10px;border-bottom:1px solid #E0DAD0">科目</th>
            <th style="text-align:center;font-size:10px;border-bottom:1px solid #E0DAD0;min-width:36px">分數</th>
            <th style="text-align:center;font-size:10px;border-bottom:1px solid #E0DAD0">等第</th>
            <th style="text-align:center;font-size:10px;border-bottom:1px solid #E0DAD0">進退步</th>
            <th style="text-align:center;font-size:10px;border-bottom:1px solid #E0DAD0">班平均</th>
            <th style="text-align:center;font-size:10px;border-bottom:1px solid #E0DAD0">差距</th>
          </tr></thead>
          <tbody>${subRows}</tbody>
        </table>
      </div>`;
  });

  // 各科平均總表
  let subAvgRows = subAvgs.map(({sub,avg,count}) => {
    const bc = avg >= 80 ? "#2E5A1A" : avg < 60 ? "#8B2222" : "#1C1A14";
    const lvl = avg >= 90 ? "優秀" : avg >= 80 ? "良好" : avg >= 70 ? "中等" : avg >= 60 ? "及格" : "待加強";
    return `<tr>
      <td style="font-size:11px;padding:3px 6px">${sub}</td>
      <td style="text-align:center;font-family:monospace;font-weight:700;font-size:12px;color:${bc}">${avg.toFixed(1)}</td>
      <td style="text-align:center;font-size:10px">${lvl}</td>
      <td style="text-align:center;font-size:10px;color:#9E9890">${count} 次</td>
    </tr>`;
  }).join("");

  // ── SVG 雷達圖（個人 + 班級平均，純 SVG 可列印）──────────
  function makePrintRadarSVG(subs, personalData, classData) {
    if (subs.length < 3) return "";
    const W = 220, H = 220, cx = 110, cy = 115, R = 80;
    const n = subs.length;
    const angles = subs.map((_, i) => (i / n) * 2 * Math.PI - Math.PI / 2);

    function polar(val, r) {
      return angles.map((a, i) => {
        const v = (val[i] ?? 0) / 100;
        return [cx + r * v * Math.cos(a), cy + r * v * Math.sin(a)];
      });
    }

    // 背景網格（20/40/60/80/100）
    let grid = "";
    [20, 40, 60, 80, 100].forEach(pct => {
      const pts = angles.map(a => [
        cx + R * (pct/100) * Math.cos(a),
        cy + R * (pct/100) * Math.sin(a)
      ]);
      grid += `<polygon points="${pts.map(p=>p.join(",")).join(" ")}" fill="none" stroke="#E0DAD0" stroke-width="0.8"/>`;
    });

    // 輻射線
    const spokes = angles.map(a =>
      `<line x1="${cx}" y1="${cy}" x2="${cx + R*Math.cos(a)}" y2="${cy + R*Math.sin(a)}" stroke="#E0DAD0" stroke-width="0.8"/>`
    ).join("");

    // 個人資料面
    const pPts = polar(personalData, R);
    const pPath = pPts.map((p,i)=>(i===0?"M":"L")+p[0].toFixed(1)+","+p[1].toFixed(1)).join(" ")+"Z";

    // 班級平均面
    let classPath = "";
    if (classData) {
      const cPts = polar(classData, R);
      classPath = `<path d="${cPts.map((p,i)=>(i===0?"M":"L")+p[0].toFixed(1)+","+p[1].toFixed(1)).join(" ")+"Z"}" fill="#C4651A0D" stroke="#C4651A" stroke-width="1.5" stroke-dasharray="4,3"/>`;
    }

    // 個人面
    const personalPath = `<path d="${pPath}" fill="#6B4FA018" stroke="#6B4FA0" stroke-width="2"/>`;

    // 個人資料點
    const dots = pPts.map(p => `<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="3" fill="#6B4FA0"/>`).join("");

    // 標籤
    const labels = subs.map((sub, i) => {
      const a = angles[i];
      const lx = cx + (R + 16) * Math.cos(a);
      const ly = cy + (R + 16) * Math.sin(a);
      const anchor = lx < cx - 5 ? "end" : lx > cx + 5 ? "start" : "middle";
      const score = personalData[i] !== null ? personalData[i].toFixed(0) : "—";
      return `<text x="${lx.toFixed(1)}" y="${(ly-4).toFixed(1)}" text-anchor="${anchor}" font-size="9" fill="#3A3629" font-weight="600">${sub}</text>
              <text x="${lx.toFixed(1)}" y="${(ly+6).toFixed(1)}" text-anchor="${anchor}" font-size="8" fill="#6B4FA0" font-weight="700">${score}</text>`;
    }).join("");

    // 圖例
    const legend = `
      <circle cx="8" cy="${H-8}" r="4" fill="#6B4FA0"/>
      <text x="15" y="${H-4}" font-size="8" fill="#6B4FA0">個人平均</text>
      ${classData ? `<line x1="70" y1="${H-8}" x2="82" y2="${H-8}" stroke="#C4651A" stroke-width="1.5" stroke-dasharray="3,2"/>
      <text x="86" y="${H-4}" font-size="8" fill="#C4651A">班級平均</text>` : ""}`;

    return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg">
      ${grid}${spokes}${classPath}${personalPath}${dots}${labels}${legend}
    </svg>`;
  }

  // 計算雷達資料
  const radarSubs = subAvgs.map(s => s.sub);
  const radarPersonal = radarSubs.map(sub => {
    const vals = scoresArr.map(sc => sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null).filter(v=>v!==null);
    return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
  });
  const radarClass = radarSubs.map(sub => {
    const vals = examsToShow.flatMap(ex =>
      S.students.map(st2 => { const v=getScores(st2.id,ex.id)[sub]; return v!==undefined&&v!==""?parseFloat(v):null; })
    ).filter(v=>v!==null);
    return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
  });
  const radarSVG = makePrintRadarSVG(radarSubs, radarPersonal, radarClass);

  return `
    <div class="an-print-page" style="padding:20px;font-family:'Noto Sans TC',sans-serif">
      <!-- 頁首 -->
      <div style="display:flex;justify-content:space-between;align-items:flex-end;border-bottom:3px solid #1C1A14;padding-bottom:10px;margin-bottom:14px">
        <div>
          <div style="font-size:10px;color:#8B7355;letter-spacing:.1em;font-weight:500;margin-bottom:4px">${getClassName()} 個人學習分析 · ${getClassYear()}</div>
          <div style="font-size:22px;font-weight:900">${st.name}</div>
          <div style="font-size:11px;color:#6B5F4A;margin-top:2px">座號 ${st.number||"—"} ／ ${semLabel} ／ 列印日期：${date}</div>
        </div>
        <div style="text-align:right;font-size:11px;color:#6B5F4A">
          ${lastTotal !== null ? `<div style="font-size:28px;font-weight:900;color:#1C1A14;font-family:monospace">${lastTotal.toFixed(0)}</div><div>最近一次總分</div>` : ""}
        </div>
      </div>

      <!-- 摘要 -->
      <div style="background:linear-gradient(135deg,#1C3A5E,#2D5F8A);color:#EAF1F8;border-radius:8px;padding:10px 14px;margin-bottom:14px;font-size:12px;line-height:1.9">
        ${lines.join("　｜　")}
      </div>

      <!-- 走勢圖 + 雷達圖 並排 -->
      <div style="display:flex;gap:16px;margin-bottom:14px;align-items:flex-start">
        <div style="flex:1">
          <div style="font-size:9px;font-weight:700;color:#8B7355;letter-spacing:.08em;text-transform:uppercase;margin-bottom:4px">▶ 總分走勢</div>
          ${makePrintTrendSVG(validPairs)}
        </div>
        ${radarSVG ? `<div style="flex-shrink:0">
          <div style="font-size:9px;font-weight:700;color:#8B7355;letter-spacing:.08em;text-transform:uppercase;margin-bottom:4px">▶ 科目雷達圖（紫=個人，橘=班級）</div>
          ${radarSVG}
        </div>` : ""}
      </div>

      <!-- 各科長條圖 -->
      ${makePrintSubjectSVG(lastSc) ? `<div style="margin-bottom:14px">
        <div style="font-size:9px;font-weight:700;color:#8B7355;letter-spacing:.08em;text-transform:uppercase;margin-bottom:4px">▶ 最新段考各科成績（${validPairs[validPairs.length-1].ex.name}）</div>
        ${makePrintSubjectSVG(lastSc)}
      </div>` : ""}

      <!-- 雙欄：各次段考 + 各科平均 -->
      <div style="display:grid;grid-template-columns:3fr 1fr;gap:16px;align-items:start">
        <div>
          <div style="font-size:10px;font-weight:700;color:#8B7355;letter-spacing:.08em;text-transform:uppercase;margin-bottom:8px">各次段考成績</div>
          ${examRows}
        </div>
        <div>
          <div style="font-size:10px;font-weight:700;color:#8B7355;letter-spacing:.08em;text-transform:uppercase;margin-bottom:8px">${semLabel}各科平均</div>
          <table style="width:100%;border-collapse:collapse">
            <thead><tr style="background:#FAF7F0">
              <th style="padding:3px 6px;font-size:10px;text-align:left;border-bottom:1px solid #E0DAD0">科目</th>
              <th style="font-size:10px;text-align:center;border-bottom:1px solid #E0DAD0">均分</th>
              <th style="font-size:10px;text-align:center;border-bottom:1px solid #E0DAD0">等第</th>
              <th style="font-size:10px;text-align:center;border-bottom:1px solid #E0DAD0">次數</th>
            </tr></thead>
            <tbody>${subAvgRows}</tbody>
          </table>

          <!-- 簽名欄 -->
          <div style="margin-top:20px;font-size:10px;color:#6B5F4A;line-height:2.5;border-top:1px solid #E0DAD0;padding-top:10px">
            <div>導師簽名：___________</div>
            <div>家長簽名：___________</div>
            <div>日期：___________</div>
          </div>
        </div>
      </div>
    </div>`;
}

// ── canvas → img 轉換輔助（列印前替換，印完還原）──────────
function canvasToImgs(container) {
  const replacements = [];
  container.querySelectorAll("canvas").forEach(canvas => {
    try {
      const img = document.createElement("img");
      img.src = canvas.toDataURL("image/png");
      img.style.cssText = `width:${canvas.offsetWidth}px;height:${canvas.offsetHeight}px;display:block`;
      canvas.parentNode.insertBefore(img, canvas);
      canvas.style.display = "none";
      replacements.push({ canvas, img });
    } catch(e) { /* cross-origin canvas 略過 */ }
  });
  return replacements;
}
function restoreCanvases(replacements) {
  replacements.forEach(({ canvas, img }) => {
    canvas.style.display = "";
    img.remove();
  });
}

function printAnalysis() {
  const studentId = S.analysisStudentId;
  if (!studentId) { showToast("請先選擇學生"); return; }
  const analysisContent = $("analysis-content");
  if (!analysisContent || !analysisContent.children.length ||
      analysisContent.querySelector(".empty-state")) {
    showToast("請先選擇學生並等待分析載入完成"); return;
  }

  // canvas → img，讓列印引擎能看到圖表
  const pcWrap = $("pc-an-wrap");
  const replacements = canvasToImgs(pcWrap || analysisContent);

  document.body.classList.add("print-analysis");
  setTimeout(() => {
    window.print();
    setTimeout(() => {
      document.body.classList.remove("print-analysis");
      restoreCanvases(replacements);
    }, 800);
  }, 150);
}

function batchPrintAnalysis() {
  if (!S.students.length) { showToast("尚無學生資料"); return; }
  const sem = $("analysis-sem")?.value || "7上";

  // 有成績的學生清單
  let examsToCheck;
  if (GRADE_MAP[sem]) {
    examsToCheck = ACTIVE_EXAMS.filter(e => GRADE_MAP[sem].semesters.includes(e.semester));
  } else {
    examsToCheck = ACTIVE_EXAMS.filter(e => e.semester === sem);
  }
  const validStudents = S.students.filter(st =>
    examsToCheck.some(ex => getFilledCount(getScores(st.id, ex.id)) > 0)
  );
  if (!validStudents.length) { showToast("所選學期目前沒有任何學生有成績資料"); return; }

  showToast(`⏳ 正在產生 ${validStudents.length} 位學生的分析，請稍候...`);

  const printWrap = $("analysis-print-wrap");
  const origStudentId = S.analysisStudentId;
  const origSem = $("analysis-sem")?.value;
  const pcAnWrap = $("pc-an-wrap");

  // 用一個隱藏的渲染容器，逐一渲染每位學生的分析頁，截圖後累積
  let idx = 0;
  const capturedPages = []; // 每人一個 { html: string }

  function renderNext() {
    if (idx >= validStudents.length) {
      // 全部截取完畢，組合並列印
      finishBatchPrint();
      return;
    }

    const st = validStudents[idx];
    showToast(`⏳ 產生中 ${idx+1}/${validStudents.length}：${st.name}`);

    // 切換到該學生，渲染分析
    S.analysisStudentId = st.id;
    const semEl = $("analysis-sem");
    if (semEl) semEl.value = sem;

    // 同步到桌機版 select
    const pcSel = $("analysis-student");
    if (pcSel) pcSel.value = st.id;

    // 渲染
    if (GRADE_MAP[sem]) {
      renderGradeSummary(st.id, st, GRADE_MAP[sem]);
    } else {
      renderAnalysis();
    }

    // 等 chart.js 渲染完成（RAF + 延遲）
    requestAnimationFrame(() => {
      setTimeout(() => {
        const content = $("analysis-content");
        if (!content) { idx++; renderNext(); return; }

        // 把 canvas 轉成 img
        const replacements = canvasToImgs(content);

        // 複製 HTML（含已轉換的 img）
        const clone = content.cloneNode(true);

        // 還原 canvas
        restoreCanvases(replacements);

        // 把 clone 的 innerHTML 存起來
        capturedPages.push(clone.outerHTML);

        idx++;
        // 稍微延遲，避免 UI 凍結
        setTimeout(renderNext, 80);
      }, 350); // 等圖表繪製完成
    });
  }

  function finishBatchPrint() {
    // 還原原本的學生選擇
    S.analysisStudentId = origStudentId;
    const semEl = $("analysis-sem");
    if (semEl && origSem) semEl.value = origSem;
    const pcSel = $("analysis-student");
    if (pcSel) pcSel.value = origStudentId || "";
    if (origStudentId) {
      if (GRADE_MAP[origSem]) renderGradeSummary(origStudentId, S.students.find(s=>s.id===origStudentId), GRADE_MAP[origSem]);
      else renderAnalysis();
    }

    if (!capturedPages.length) { showToast("沒有可列印的資料"); return; }

    // 組合所有頁面，包上分頁 wrapper
    const allHtml = capturedPages.map((html, i) =>
      `<div style="page-break-after:${i < capturedPages.length-1 ? 'always' : 'avoid'};padding:0;margin:0">${html}</div>`
    ).join("");

    printWrap.innerHTML = allHtml;
    document.body.classList.add("print-analysis-batch");
    showToast(`🖨 準備列印 ${capturedPages.length} 位學生...`);

    setTimeout(() => {
      window.print();
      setTimeout(() => {
        document.body.classList.remove("print-analysis-batch");
        printWrap.innerHTML = "";
      }, 1000);
    }, 300);
  }

  // 開始逐一渲染
  setTimeout(renderNext, 200);
}

// ── 外觀設定頁 ────────────────────────────────────────────────
async function initSettingsEditor() {
  let cfg = { ...HP_DEFAULTS };
  if (db) {
    try {
      const doc = await col("config").doc("homepage_style").get();
      if (doc.exists) cfg = { ...HP_DEFAULTS, ...doc.data() };
    } catch(e) {}
  }
  const set = (id,val) => { const el=$(id); if(el) el.value=val; };
  set("hp-school",cfg.school); set("hp-title",cfg.title); set("hp-en",cfg.en);
  set("hp-sub",cfg.sub); set("hp-quote",cfg.quote);
  const sc=(id,v)=>{ const p=$(id),h=$(id+"-hex"); if(p)p.value=v; if(h)h.value=v; };
  sc("hp-bg-color",cfg.bgColor); sc("hp-title-color",cfg.titleColor);
  sc("hp-accent-color",cfg.accentColor); sc("hp-sub-color",cfg.subColor);
  sc("hp-right-bg",cfg.rightBg); sc("hp-body-bg",cfg.bodyBg);
  hpPreview();
}

function hpPreview() {
  const g = id => { const el=$(id); return el?el.value:""; };
  const bgClr    = g("hp-bg-color")    || HP_DEFAULTS.bgColor;
  const titleClr = g("hp-title-color") || HP_DEFAULTS.titleColor;
  const accentClr= g("hp-accent-color")|| HP_DEFAULTS.accentColor;
  const subClr   = g("hp-sub-color")   || HP_DEFAULTS.subColor;

  const bar = $("hp-preview-bar"); if(bar) bar.style.background = bgClr;
  const ps=$("prev-school"),pt=$("prev-title"),pe=$("prev-en"),psb=$("prev-sub"),pq=$("prev-quote");
  if(ps){ps.textContent=g("hp-school"); ps.style.color=subClr;}
  if(pt){pt.innerHTML=g("hp-title").replace(/\n/g,"<br>"); pt.style.color=titleClr;}
  if(pe){pe.textContent=g("hp-en"); pe.style.color=accentClr;}
  if(psb){psb.textContent=g("hp-sub"); psb.style.color=subClr;}
  if(pq){pq.textContent=g("hp-quote"); pq.style.color=subClr;}
}

async function saveHomepageStyle() {
  const g=id=>{const el=$(id);return el?el.value:"";};
  const cfg={
    school:g("hp-school"), title:g("hp-title"), en:g("hp-en"), sub:g("hp-sub"), quote:g("hp-quote"),
    bgColor:g("hp-bg-color"), titleColor:g("hp-title-color"), accentColor:g("hp-accent-color"),
    subColor:g("hp-sub-color"), rightBg:g("hp-right-bg"), bodyBg:g("hp-body-bg")
  };
  if(db){ try { await col("config").doc("homepage_style").set(cfg); } catch(e){ showToast("儲存失敗："+e.message); return; } }
  applyLoginStyle(cfg);
  show("hp-save-msg"); setTimeout(()=>hide("hp-save-msg"),3000);
  showToast("✅ 首頁設定已儲存並生效！");
}

function applyTheme(name) {
  const t = HP_THEMES[name]; if(!t) return;
  const sc=(id,v)=>{const p=$(id),h=$(id+"-hex");if(p)p.value=v;if(h)h.value=v;};
  sc("hp-bg-color",t.bgColor); sc("hp-title-color",t.titleColor);
  sc("hp-accent-color",t.accentColor); sc("hp-sub-color",t.subColor);
  sc("hp-right-bg",t.rightBg); sc("hp-body-bg",t.bodyBg);
  hpPreview(); showToast("✨ 主題已套用，記得點「儲存設定」");
}

async function resetHomepageStyle() {
  if (!confirm("確定要恢復所有首頁設定為預設值嗎？")) return;
  if (db) { col("config").doc("homepage_style").delete().catch(()=>{}); }
  applyLoginStyle(HP_DEFAULTS);
  initSettingsEditor();
  showToast("↺ 已恢復預設設定");
}

async function loadLoginStyle() {
  if (!db) return;
  try {
    const doc = await col("config").doc("homepage_style").get();
    if (doc.exists) applyLoginStyle({ ...HP_DEFAULTS, ...doc.data() });
  } catch(e) {}
}

function applyLoginStyle(cfg) {
  const s = { ...HP_DEFAULTS, ...cfg };
  const qs = sel => document.querySelector(sel);
  const schoolEl = qs(".masthead-name");
  if (schoolEl) { schoolEl.textContent = s.school; schoolEl.style.color = s.subColor; }
  const titleEl = qs(".masthead-title");
  if (titleEl) { titleEl.innerHTML = s.title.replace(/\n/g,"<br>"); titleEl.style.color = s.titleColor; }
  const enEl = qs(".masthead-en");
  if (enEl) { enEl.textContent = s.en; enEl.style.color = s.accentColor; }
  const subEl = qs(".masthead-sub");
  if (subEl) { subEl.textContent = s.sub; subEl.style.color = s.subColor; }
  const quoteEl = qs(".pull-quote p");
  if (quoteEl) quoteEl.textContent = s.quote;
  const loginLeft = qs(".login-left");
  if (loginLeft) loginLeft.style.background = s.bgColor;
  const loginRight = qs(".login-right");
  if (loginRight) loginRight.style.background = s.rightBg;
  document.body.style.background = s.bodyBg;
}

if (db) loadLoginStyle();

// ── 資料匯出 / 匯入 ───────────────────────────────────────────
function exportData() {
  const payload = {
    students: S.students,
    scores:   S.scores,
    teacherComments: S.teacherComments,
    studentMemos:    S.studentMemos,
    exportedAt: new Date().toISOString(),
    version: "113-3-v2"
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type:"application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `113-3班完整備份_${new Date().toISOString().slice(0,10)}.json`;
  a.click();
}

function importData(event) {
  const file = event.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = async e => {
    try {
      const data = JSON.parse(e.target.result);
      if (!data.students) { showToast("檔案格式不正確"); return; }
      if (!confirm("匯入將合併現有資料，是否繼續？")) return;
      // 合併學生（以 id 去重）
      data.students.forEach(st => {
        if (!S.students.find(s=>s.id===st.id)) S.students.push(st);
      });
      Object.assign(S.scores, data.scores||{});
      // 寫回 Firestore
      if (db) {
        const batch = db.batch();
        data.students.forEach(st => batch.set(col("students").doc(st.id), { name:st.name, number:st.number||0 }));
        Object.entries(data.scores||{}).forEach(([key,val]) => batch.set(col("scores").doc(key), val));
        await batch.commit();
      }
      saveLocalData(); updateHeaderCount(); renderStudentList();
      showToast("✅ 匯入成功！");
    } catch(e) { showToast("檔案解析失敗：" + e.message); }
  };
  reader.readAsText(file);
  event.target.value = "";
}


// ── 下載範例 Excel ────────────────────────────────────────────
function downloadSampleExcel() {
  if (typeof XLSX === "undefined") { showToast("⚠️ Excel 套件尚未載入，請稍後再試"); return; }

  const subjects = ACTIVE_SUBJECTS.length ? ACTIVE_SUBJECTS : ["國文","英語文","數學","生物","理化","地科","歷史","地理","公民"];
  const exams    = ACTIVE_EXAMS.length    ? ACTIVE_EXAMS    : [];
  const wb = XLSX.utils.book_new();

  // ── 說明文字 ──────────────────────────────────────────────
  const infoRows = [
    ["📋 成績追蹤系統 — Excel 匯入範例檔"],
    [`班級：${getClassName()}　學年：${getClassYear()}　產生日期：${new Date().toLocaleDateString("zh-TW")}`],
    [],
    ["【使用說明】"],
    ["1. 本檔包含「學生名單」工作表及各次段考工作表（共 " + (1 + exams.length) + " 張）"],
    ["2. 先填寫「學生名單」：座號、姓名（每行一位）"],
    ["3. 再填各段考工作表：座號、姓名、各科分數（0–100）、校排（選填）"],
    ["4. 班排系統匯入後自動計算，無需填寫"],
    ["5. 缺考或未應試的格子請留空，不要填 0"],
    ["6. 存檔時請維持 .xlsx 格式"],
    [],
    [`科目欄位順序（共 ${subjects.length} 科）：` + subjects.join("、")],
  ];
  const infoWs = XLSX.utils.aoa_to_sheet(infoRows);
  infoWs["!cols"] = [{ wch: 80 }];
  XLSX.utils.book_append_sheet(wb, infoWs, "說明");

  // ── 學生名單 sheet ─────────────────────────────────────────
  const nameHeader1 = ["學生名單範例"];
  const nameHeader2 = ["座號", "姓名", "（座號與姓名為必填；有缺號請整列留空）"];
  const nameColNames = ["座號", "姓名"];
  const sampleNames = [
    [1,"王小明"],[2,"李小華"],[3,"張美玲"],[4,""],[5,"陳大為"],
    [6,"林志玲"],[7,"吳宗憲"],[8,"黃安琪"],[9,"劉建志"],[10,"楊雅婷"],
  ];
  const nameData = [nameHeader1, nameHeader2, nameColNames, ...sampleNames];
  const nameWs = XLSX.utils.aoa_to_sheet(nameData);
  nameWs["!cols"] = [{ wch: 8 }, { wch: 12 }, { wch: 40 }];
  // 合併標題列
  nameWs["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:2} }];
  XLSX.utils.book_append_sheet(wb, nameWs, "學生名單");

  // ── 各段考 sheet ───────────────────────────────────────────
  const examColWidth = [{ wch: 8 }, { wch: 10 }, ...subjects.map(() => ({ wch: 7 })), { wch: 8 }];

  exams.forEach(exam => {
    const header1  = [exam.name];
    const header2  = ["座號、姓名為必填；各科分數 0–100，留空表示缺考；班排系統自動計算無需填寫"];
    const colNames = ["座號", "姓名", ...subjects, "校排"];
    const rows = [];
    for (let i = 1; i <= 10; i++) {
      // 第 4 號示範缺號留空
      if (i === 4) { rows.push([]); continue; }
      rows.push([i, "", ...subjects.map(() => ""), ""]);
    }
    const sheetData = [header1, header2, colNames, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws["!cols"] = examColWidth;
    ws["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c: subjects.length + 2} }];
    XLSX.utils.book_append_sheet(wb, ws, exam.name);
  });

  // ── 若尚無段考設定，產生一張示範 sheet ────────────────────
  if (exams.length === 0) {
    const colNames = ["座號", "姓名", ...subjects, "校排"];
    const rows = Array.from({ length: 10 }, (_, i) => [i + 1, "", ...subjects.map(() => ""), ""]);
    const ws = XLSX.utils.aoa_to_sheet([["（示範）7上第一次段考"], ["座號、姓名為必填"], colNames, ...rows]);
    ws["!cols"] = examColWidth;
    ws["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c: subjects.length + 2} }];
    XLSX.utils.book_append_sheet(wb, ws, "7上第一次段考");
  }

  const className = getClassName().replace(/\s/g, "_");
  XLSX.writeFile(wb, `成績匯入範例_${className}.xlsx`);
  showToast("✅ 範例檔已下載！");
}


// ── Excel 匯入 ───────────────────────────────────────────────
function importExcel(event) {
  const file = event.target.files[0]; if (!file) return;
  if (typeof XLSX === "undefined") { showToast("⚠️ Excel 套件尚未載入，請稍後再試"); return; }

  const reader = new FileReader();
  reader.onload = async e => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });
      let importedStudents = 0, importedScores = 0, errors = [];

      // 輔助：讀取 sheet 並跳過前 3 列（標題、說明、欄位名稱），從第 4 列取資料
      function readSheetRows(sheet) {
        const all = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        return all.slice(3); // 跳過列0(標題)、列1(說明)、列2(欄位名稱)
      }

      // ── Sheet「學生名單」──────────────────────────────────
      const nameSheet = wb.Sheets["學生名單"];
      let skippedNumbers = []; // 記錄跳過的缺號座號
      if (nameSheet) {
        const rows = readSheetRows(nameSheet);
        rows.forEach(row => {
          const number = parseInt(row[0]);
          const name   = (row[1] || "").toString().trim();
          // 座號有填但姓名空白 → 缺號，記錄並跳過
          if (number && !name) { skippedNumbers.push(number); return; }
          // 座號也沒填 → 完全空白列，直接跳過
          if (!number || !name) return;
          // 檢查座號是否重複
          const existing = S.students.find(s => parseInt(s.number) === number);
          if (existing) {
            existing.name = name; // 更新姓名
          } else {
            const id = "st_" + number.toString().padStart(2,"0") + "_" + Date.now().toString(36);
            S.students.push({ id, name, number });
            importedStudents++;
          }
        });
        S.students.sort((a,b) => (parseInt(a.number)||999) - (parseInt(b.number)||999));
      }

      // ── 段考 Sheet（依 ACTIVE_EXAMS 設定逐一比對）──────────────
      const examMap = {};
      ACTIVE_EXAMS.forEach(e => { examMap[e.name] = e.id; });

      for (const [sheetName, examId] of Object.entries(examMap)) {
        const ws = wb.Sheets[sheetName];
        if (!ws) continue;
        const rows = readSheetRows(ws);
        rows.forEach(row => {
          const number = parseInt(row[0]);
          const name   = (row[1] || "").toString().trim();
          // 座號或姓名皆空 → 缺號列，跳過
          if (!number || !name) return;
          // 優先用座號比對，找不到才試姓名
          const st = S.students.find(s => parseInt(s.number) === number)
                  || S.students.find(s => s.name === name);
          if (!st) return; // 缺號或名單中不存在的座號，靜默跳過
          const sc = {};
          ACTIVE_SUBJECTS.forEach((sub, i) => {
            const val = row[2 + i];
            if (val !== undefined && val !== "" && val !== null) {
              const n = parseFloat(val);
              if (!isNaN(n) && n >= 0 && n <= 100) sc[sub] = n;
            }
          });
          // 班排由系統自動計算，Excel 只有校排欄
          const crVal = row[2 + ACTIVE_SUBJECTS.length];
          if (crVal !== "" && !isNaN(parseInt(crVal))) sc["校排"] = parseInt(crVal);
          if (Object.keys(sc).length === 0) return; // 這列沒有任何成績，跳過
          sc["_updatedAt"] = new Date().toISOString();
          const key = scoreKey(st.id, examId);
          S.scores[key] = { ...(S.scores[key]||{}), ...sc };
          importedScores++;
        });
      }

      // ── 寫回 Firestore ───────────────────────────────────
      if (db) {
        const batch = db.batch();
        S.students.forEach(st =>
          batch.set(col("students").doc(st.id), { name:st.name, number:st.number||0 })
        );
        Object.entries(S.scores).forEach(([key,val]) =>
          batch.set(col("scores").doc(key), val)
        );
        await batch.commit();
      }
      // 匯入後自動計算各次段考的班排
      ACTIVE_EXAMS.forEach(ex => calcClassRanks(ex.id));
      // 把更新後的班排也寫回 Firestore
      if (db) {
        const batch2 = db.batch();
        S.students.forEach(st => {
          ACTIVE_EXAMS.forEach(ex => {
            const key = scoreKey(st.id, ex.id);
            if (S.scores[key]) batch2.set(col("scores").doc(key), S.scores[key]);
          });
        });
        await batch2.commit().catch(e => console.warn("班排寫回失敗:", e));
      }
      saveLocalData(); updateHeaderCount(); renderStudentList();
      const skipMsg = skippedNumbers.length > 0 ? `\n⚠️ 自動跳過缺號座號：${skippedNumbers.join("、")} 號` : "";
      showToast(`✅ 匯入完成！新增 ${importedStudents} 位學生、更新 ${importedScores} 筆成績（班排已自動計算）${skipMsg}`);
      if (skippedNumbers.length > 0) console.log("匯入跳過缺號：", skippedNumbers);

    } catch(err) {
      showToast("❌ Excel 解析失敗：" + err.message);
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
  event.target.value = "";
}

// ══════════════════════════════════════════════════════════
// 功能 6：全班成績熱力圖
// ══════════════════════════════════════════════════════════
function renderHeatmap(examId) {
  const wrap = $("heatmap-wrap"); if (!wrap) return;
  if (!S.students.length) { wrap.innerHTML="<div style='font-size:12px;color:#9E9890;padding:8px'>尚無學生資料</div>"; return; }
  const data = S.students.map(st => ({
    st, scores: ACTIVE_SUBJECTS.map(sub => { const v=getScores(st.id,examId)[sub]; return v!==undefined&&v!==""?parseFloat(v):null; })
  }));
  const hasAny = data.some(d=>d.scores.some(v=>v!==null));
  if (!hasAny) { wrap.innerHTML="<div style='font-size:12px;color:#9E9890;padding:8px'>本次段考尚無資料</div>"; return; }

  function scoreColor(val) {
    if (val===null) return "#F5F0E8";
    if (val>=90) return "#1A5C2A"; if (val>=80) return "#2E7A3A";
    if (val>=70) return "#5B8A4A"; if (val>=60) return "#C4651A";
    if (val>=50) return "#D4822A"; return "#A83232";
  }

  const sorted = [...data].sort((a,b)=>{
    const ta=a.scores.filter(v=>v!==null).reduce((s,v)=>s+v,0);
    const tb=b.scores.filter(v=>v!==null).reduce((s,v)=>s+v,0);
    return tb-ta;
  });

  let html=`<div style="overflow-x:auto;-webkit-overflow-scrolling:touch"><table style="border-collapse:separate;border-spacing:2px;min-width:600px">
    <thead><tr><th style="padding:6px 8px;text-align:left;font-size:11px;color:#6B5F4A;font-weight:600">學生</th>`;
  ACTIVE_SUBJECTS.forEach(s=>{ html+=`<th style="padding:6px 4px;text-align:center;font-size:11px;color:#6B5F4A;font-weight:600;white-space:nowrap">${s}</th>`; });
  html+=`<th style="padding:6px 8px;text-align:center;font-size:11px;color:#6B5F4A;font-weight:600">總分</th></tr></thead><tbody>`;

  sorted.forEach(({st,scores})=>{
    const total=scores.filter(v=>v!==null).reduce((s,v)=>s+v,0);
    const has=scores.some(v=>v!==null);
    html+=`<tr><td style="padding:4px 8px;font-size:12px;font-weight:600;white-space:nowrap;color:#1C1A14">${st.name}</td>`;
    scores.forEach(val=>{
      html+=`<td style="padding:2px;text-align:center"><div style="background:${scoreColor(val)};color:#fff;border-radius:4px;padding:5px 3px;font-size:11px;font-family:'DM Mono',monospace;font-weight:600;min-width:32px">${val!==null?val.toFixed(0):"·"}</div></td>`;
    });
    const hmMax=has?getStudentMaxScore(st.id,examId):ACTIVE_SUBJECTS.length*100;
    html+=`<td style="padding:4px 6px;text-align:center;font-family:'DM Mono',monospace;font-weight:700;font-size:13px;color:${has?(total>=hmMax*0.7?"#2E5A1A":total<hmMax*0.6?"#8B2222":"#1C1A14"):"#C8BA9E"}">${has?total.toFixed(0):"—"}</td></tr>`;
  });

  // 班級平均列
  html+=`<tr style="border-top:2px solid #C8BA9E;background:#FAF7F0"><td style="padding:6px 8px;font-size:11px;font-weight:700;color:#6B5F4A">班級平均</td>`;
  ACTIVE_SUBJECTS.forEach((sub,si)=>{
    const vals=data.map(d=>d.scores[si]).filter(v=>v!==null);
    const av=vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null;
    const bg=av!==null?(av>=80?"#5B8A4A":av<60?"#A83232":"#2D5F8A"):"#E0DAD0";
    html+=`<td style="padding:2px;text-align:center"><div style="background:${bg};color:#fff;border-radius:4px;padding:5px 3px;font-size:11px;font-family:'DM Mono',monospace;font-weight:700;min-width:32px">${av!==null?av.toFixed(1):"·"}</div></td>`;
  });
  const allT=data.map(d=>d.scores.filter(v=>v!==null).reduce((s,v)=>s+v,0)).filter((_,i)=>data[i].scores.some(v=>v!==null));
  const cT=allT.length?allT.reduce((a,b)=>a+b,0)/allT.length:null;
  html+=`<td style="padding:4px 6px;text-align:center;font-family:'DM Mono',monospace;font-weight:700;font-size:13px;color:#2D5F8A">${cT!==null?cT.toFixed(1):"—"}</td></tr>`;
  html+=`</tbody></table></div>
  <div style="display:flex;gap:12px;flex-wrap:wrap;margin-top:10px;font-size:11px;color:#6B5F4A">
    <span>色階：</span>
    <span style="display:flex;align-items:center;gap:4px"><span style="background:#1A5C2A;color:#fff;padding:2px 6px;border-radius:3px;font-family:monospace">90+</span>優秀</span>
    <span style="display:flex;align-items:center;gap:4px"><span style="background:#5B8A4A;color:#fff;padding:2px 6px;border-radius:3px;font-family:monospace">70-89</span>良好</span>
    <span style="display:flex;align-items:center;gap:4px"><span style="background:#C4651A;color:#fff;padding:2px 6px;border-radius:3px;font-family:monospace">60-69</span>及格</span>
    <span style="display:flex;align-items:center;gap:4px"><span style="background:#A83232;color:#fff;padding:2px 6px;border-radius:3px;font-family:monospace">&lt;60</span>待加強</span>
  </div>`;
  wrap.innerHTML=html;
}

// ══════════════════════════════════════════════════════════
// 功能 4+5：全學期總結頁
// ══════════════════════════════════════════════════════════
function renderSummaryPage() {
  buildExamOptions("summary-exam", "", false, "");
  const se=$("summary-exam");
  if(se) se.innerHTML='<option value="">— 選擇段考（相關性分析）—</option>'+se.innerHTML;
  renderAllSemesterTrend();
  renderSubjectLongTrend();
  renderSchoolRankTrend();
  renderMilestones();
  renderCorrelation();
  renderSemesterSummaryTable();
  renderGrowthRanking();
  renderGroupAnalysis();
  loadNotes();
}

function renderAllSemesterTrend() {
  destroyChart("chart-all-sem-trend");
  const ctx=$("chart-all-sem-trend"); if(!ctx) return;
  const classAvgs=ACTIVE_EXAMS.map(ex=>{const t=S.students.map(st=>getTotal(getScores(st.id,ex.id))).filter(v=>v!==null);return t.length?t.reduce((a,b)=>a+b,0)/t.length:null;});
  const highs=ACTIVE_EXAMS.map(ex=>{const t=S.students.map(st=>getTotal(getScores(st.id,ex.id))).filter(v=>v!==null);return t.length?Math.max(...t):null;});
  const lows=ACTIVE_EXAMS.map(ex=>{const t=S.students.map(st=>getTotal(getScores(st.id,ex.id))).filter(v=>v!==null);return t.length?Math.min(...t):null;});
  chartInstances["chart-all-sem-trend"]=new Chart(ctx,{
    type:"line",
    data:{labels:ACTIVE_EXAMS.map(e=>e.name.replace("次段考","").replace("第","").trim()),
      datasets:[
        {label:"班級平均",data:classAvgs,borderColor:"#2D5F8A",backgroundColor:"#2D5F8A15",borderWidth:2.5,pointRadius:4,fill:true,tension:0.3,spanGaps:true},
        {label:"最高總分",data:highs,borderColor:"#5B8A4A",backgroundColor:"transparent",borderWidth:1.5,pointRadius:3,borderDash:[4,3],tension:0.3,spanGaps:true},
        {label:"最低總分",data:lows,borderColor:"#A83232",backgroundColor:"transparent",borderWidth:1.5,pointRadius:3,borderDash:[4,3],tension:0.3,spanGaps:true},
      ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:"bottom",labels:{font:{size:11},boxWidth:12}}},
      scales:{y:{min:0,max:ACTIVE_SUBJECTS.length*100,grid:{color:"#E2DED6"},ticks:{font:{size:11}}},x:{grid:{color:"#F0EAE0"},ticks:{font:{size:9},maxRotation:45}}}}
  });
}

function renderMilestones() {
  const wrap=$("milestones-wrap"); if(!wrap) return;
  const milestones=[];
  SEMESTERS.forEach(sem=>{
    const semExams=ACTIVE_EXAMS.filter(e=>e.semester===sem);
    const examData=semExams.map(ex=>{const t=S.students.map(st=>getTotal(getScores(st.id,ex.id))).filter(v=>v!==null);return{ex,totals:t};}).filter(d=>d.totals.length>0);
    if(examData.length<2) return;
    const first=examData[0],last=examData[examData.length-1];
    const firstAvg=first.totals.reduce((a,b)=>a+b,0)/first.totals.length;
    const lastAvg=last.totals.reduce((a,b)=>a+b,0)/last.totals.length;
    const diff=lastAvg-firstAvg;
    if(Math.abs(diff)>=5) milestones.push({sem,text:diff>0?`${sem}學期班級平均進步 ${diff.toFixed(1)} 分（${firstAvg.toFixed(1)} → ${lastAvg.toFixed(1)}）`:`${sem}學期班級平均退步 ${Math.abs(diff).toFixed(1)} 分（${firstAvg.toFixed(1)} → ${lastAvg.toFixed(1)}）`,type:diff>0?"up":"down"});
    const msMax=getExamMaxScore(last.ex.id);
    const msPass=msMax*0.6;
    const passRate=last.totals.filter(v=>v>=msPass).length/last.totals.length*100;
    if(passRate>=80) milestones.push({sem,text:`${sem}學期末及格率達 ${passRate.toFixed(0)}%（滿分${msMax}分）`,type:"star"});
    else if(passRate<50) milestones.push({sem,text:`${sem}學期末及格率僅 ${passRate.toFixed(0)}%，需加強`,type:"warn"});
  });
  if(!milestones.length){wrap.innerHTML='<div style="font-size:12px;color:#9E9890;padding:8px">需要至少兩個學期的資料才能顯示里程碑</div>';return;}
  const icons={up:"📈",down:"📉",star:"⭐",warn:"⚠️"};
  const colors={up:"#EDF4EA",down:"#FAECEC",star:"#FEF7E0",warn:"#FDF0E6"};
  const tcs={up:"#2E5A1A",down:"#8B2222",star:"#8B6A00",warn:"#8B4A14"};
  wrap.innerHTML=milestones.map(m=>`<div style="display:flex;align-items:flex-start;gap:10px;padding:10px 0;border-bottom:1px solid #F0EAE0"><span style="font-size:18px;flex-shrink:0">${icons[m.type]}</span><div><span style="display:inline-block;font-size:11px;font-weight:600;padding:1px 7px;border-radius:99px;background:${colors[m.type]};color:${tcs[m.type]};margin-bottom:4px">${m.sem}</span><div style="font-size:13px;color:#1C1A14">${m.text}</div></div></div>`).join("");
}

function renderCorrelation() {
  const wrap=$("correlation-wrap"); if(!wrap) return;
  const examId=$("summary-exam")?.value;
  if(!examId){wrap.innerHTML='<div style="font-size:12px;color:#9E9890;padding:8px">請從右上角選擇段考</div>';return;}
  function pearson(xs,ys){
    const n=xs.length; if(n<3) return null;
    const mx=xs.reduce((a,b)=>a+b,0)/n,my=ys.reduce((a,b)=>a+b,0)/n;
    const num=xs.reduce((s,x,i)=>s+(x-mx)*(ys[i]-my),0);
    const dx=Math.sqrt(xs.reduce((s,x)=>s+Math.pow(x-mx,2),0));
    const dy=Math.sqrt(ys.reduce((s,y)=>s+Math.pow(y-my,2),0));
    return(dx===0||dy===0)?null:num/(dx*dy);
  }
  const pairs=[];
  for(let i=0;i<ACTIVE_SUBJECTS.length;i++){
    for(let j=i+1;j<ACTIVE_SUBJECTS.length;j++){
      const combined=S.students.map(st=>{const sc=getScores(st.id,examId);const vi=sc[ACTIVE_SUBJECTS[i]],vj=sc[ACTIVE_SUBJECTS[j]];return(vi!==undefined&&vi!==""&&vj!==undefined&&vj!=="")?[parseFloat(vi),parseFloat(vj)]:null;}).filter(Boolean);
      if(combined.length<3) continue;
      const r=pearson(combined.map(c=>c[0]),combined.map(c=>c[1]));
      if(r!==null) pairs.push({a:ACTIVE_SUBJECTS[i],b:ACTIVE_SUBJECTS[j],r,n:combined.length});
    }
  }
  pairs.sort((a,b)=>Math.abs(b.r)-Math.abs(a.r));
  const top=pairs.slice(0,6);
  if(!top.length){wrap.innerHTML='<div style="font-size:12px;color:#9E9890;padding:8px">資料不足，無法計算相關性</div>';return;}

  // 樣本數警示
  const sampleN = S.students.length;
  const sampleWarn = sampleN < 20
    ? `<div style="font-size:11px;color:#C4651A;background:#FDF0E6;border:1px solid #F0C090;border-radius:6px;padding:8px 12px;margin-bottom:10px">⚠️ 目前班級樣本數為 ${sampleN} 人，樣本較小時相關係數可能不穩定，結果僅供參考，請勿過度解讀。</div>`
    : "";

  wrap.innerHTML= sampleWarn + top.map(({a,b,r,n})=>{
    const pct=Math.abs(r)*100;
    const isPos=r>0;
    const strength=Math.abs(r)>=0.7?"強":Math.abs(r)>=0.4?"中":"弱";
    // 簡易顯著性估算：r 的 t 值 = r*sqrt(n-2)/sqrt(1-r^2)，df=n-2
    // p<.05 的臨界 t 值：n=10→2.31, n=15→2.16, n=20→2.09, n=30→2.05
    const t = Math.abs(r)*Math.sqrt(n-2)/Math.sqrt(1-r*r+1e-10);
    const critT = n>=30?2.05:n>=20?2.09:n>=15?2.16:n>=10?2.31:999;
    const isSig = t >= critT;
    const bc=isPos?"#2D5F8A":"#C4651A";
    const desc=isPos?`${a}高分的學生，${b}通常也較高`:`${a}高分的學生，${b}通常較低`;
    const sigBadge = isSig
      ? `<span style="font-size:10px;padding:1px 6px;border-radius:99px;background:#EDF4EA;color:#2E5A1A;font-weight:600">p&lt;.05 顯著</span>`
      : `<span style="font-size:10px;padding:1px 6px;border-radius:99px;background:#F0EAE0;color:#9E9890;font-weight:600">不顯著</span>`;
    return `<div style="margin-bottom:10px;padding:12px;background:#FAF7F0;border-radius:8px;border:1px solid #E0DAD0"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px"><div style="font-size:13px;font-weight:600;color:#1C1A14">${a} × ${b}</div><div style="display:flex;gap:6px;align-items:center">${sigBadge}<span style="font-size:11px;padding:2px 7px;border-radius:99px;background:${bc}22;color:${bc};font-weight:600">${strength}相關</span><span style="font-family:'DM Mono',monospace;font-size:13px;font-weight:700;color:${bc}">${isPos?"+":""}${r.toFixed(2)}</span></div></div><div style="height:6px;background:#E0DAD0;border-radius:3px;overflow:hidden;margin-bottom:6px"><div style="width:${pct.toFixed(1)}%;background:${bc};height:100%;border-radius:3px"></div></div><div style="font-size:11px;color:#6B5F4A">${desc}（${n} 位學生）</div></div>`;
  }).join("");
}

function renderSemesterSummaryTable() {
  const wrap=$("semester-summary-table"); if(!wrap) return;
  let html=`<div class="table-wrap"><table><thead><tr><th>學期</th><th style="text-align:center">資料次數</th><th style="text-align:center">最新班平均</th><th style="text-align:center">最高</th><th style="text-align:center">最低</th><th style="text-align:center">及格率</th><th style="text-align:center">學期趨勢</th></tr></thead><tbody>`;
  SEMESTERS.forEach(sem=>{
    const semExams=ACTIVE_EXAMS.filter(e=>e.semester===sem);
    const examData=semExams.map(ex=>{const t=S.students.map(st=>getTotal(getScores(st.id,ex.id))).filter(v=>v!==null);return{ex,totals:t};}).filter(d=>d.totals.length>0);
    if(!examData.length){html+=`<tr><td style="font-weight:600">${sem}</td><td colspan="6" style="text-align:center;color:#C8BA9E;font-size:12px">尚無資料</td></tr>`;return;}
    const latest=examData[examData.length-1],first=examData[0];
    const latestAvg=latest.totals.reduce((a,b)=>a+b,0)/latest.totals.length;
    const firstAvg=first.totals.reduce((a,b)=>a+b,0)/first.totals.length;
    const diff=examData.length>=2?latestAvg-firstAvg:null;
    const stMax=getExamMaxScore(latest.ex.id);
    const passLine=stMax*0.6;
    const passRate=latest.totals.filter(v=>v>=passLine).length/latest.totals.length*100;
    const semMax=getExamMaxScore(latest.ex.id);
    const ac=latestAvg>=semMax*0.7?"#2E5A1A":latestAvg<semMax*0.6?"#8B2222":"#1C1A14";
    const trendHtml=diff===null?"—":diff>0?`<span style="color:#2E5A1A;font-weight:600">▲ ${diff.toFixed(1)}</span>`:`<span style="color:#8B2222;font-weight:600">▼ ${Math.abs(diff).toFixed(1)}</span>`;
    html+=`<tr><td style="font-weight:600">${sem}</td><td style="text-align:center;color:#9E9890">${examData.length}/${semExams.length}</td><td style="text-align:center;font-family:'DM Mono',monospace;font-weight:700;color:${ac}">${latestAvg.toFixed(1)}</td><td style="text-align:center;font-family:'DM Mono',monospace;color:#2E5A1A">${Math.max(...latest.totals).toFixed(0)}</td><td style="text-align:center;font-family:'DM Mono',monospace;color:#8B2222">${Math.min(...latest.totals).toFixed(0)}</td><td style="text-align:center;color:${passRate>=70?"#2E5A1A":passRate>=50?"#C4651A":"#8B2222"};font-weight:600">${passRate.toFixed(0)}%</td><td style="text-align:center">${trendHtml}</td></tr>`;
  });
  html+=`</tbody></table></div>`;
  wrap.innerHTML=html;
}

// ══════════════════════════════════════════════════════════
// 功能 1：全班學生成長排行榜
// ══════════════════════════════════════════════════════════
function renderGrowthRanking() {
  const wrap = $("growth-ranking-wrap"); if (!wrap) return;
  const period = $("growth-period-sel")?.value || "all";

  function linRegSlope(vals) {
    const n = vals.length; if (n < 2) return null;
    const xs = vals.map((_,i) => i);
    const mx = (n-1)/2;
    const my = vals.reduce((a,b)=>a+b,0)/n;
    const num = xs.reduce((s,x,i) => s + (x-mx)*(vals[i]-my), 0);
    const den = xs.reduce((s,x)   => s + (x-mx)**2, 0);
    return den === 0 ? 0 : num/den;
  }

  // 依時間範圍過濾段考
  function getExamsForPeriod() {
    if (period === "all")     return ACTIVE_EXAMS;
    if (period === "recent2") return null; // 動態取每人最近2次
    if (period === "recent3") return null;
    return ACTIVE_EXAMS.filter(e => e.semester === period);
  }
  const periodExams = getExamsForPeriod();

  const rows = S.students.map((st, idx) => {
    // 取該學生在指定範圍內的成績
    let validPairs;
    if (period === "recent2" || period === "recent3") {
      const n = period === "recent2" ? 2 : 3;
      const all = ACTIVE_EXAMS.map(ex => ({ ex, t: getTotal(getScores(st.id, ex.id)) })).filter(x=>x.t!==null);
      validPairs = all.slice(-n);
    } else {
      const exams = periodExams || ACTIVE_EXAMS;
      validPairs = exams.map(ex => ({ ex, t: getTotal(getScores(st.id, ex.id)) })).filter(x=>x.t!==null);
    }
    if (validPairs.length < 2) return null;
    const slope    = linRegSlope(validPairs.map(v=>v.t));
    const rawDiff  = validPairs[validPairs.length-1].t - validPairs[0].t;
    const firstEx  = validPairs[0].ex, lastEx = validPairs[validPairs.length-1].ex;
    const firstT   = validPairs[0].t,  lastT  = validPairs[validPairs.length-1].t;
    return { st, idx, slope, rawDiff, firstT, lastT, firstEx, lastEx, dataCount: validPairs.length };
  }).filter(Boolean);

  if (!rows.length) {
    wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">需要至少兩次段考資料才能計算成長</div>';
    return;
  }

  rows.sort((a,b) => b.slope - a.slope);
  const RC = ["#2D5F8A","#5B8A4A","#C4651A","#6B4FA0","#A83232"];
  const RL = ["#EAF1F8","#EDF4EA","#FDF0E6","#F2EEF8","#FAECEC"];

  // 斜率的說明標籤
  function slopeLabel(slope) {
    if (slope > 15)  return "急速進步";
    if (slope > 5)   return "持續進步";
    if (slope > 0)   return "緩步上升";
    if (slope > -5)  return "小幅下滑";
    if (slope > -15) return "持續退步";
    return "急速退步";
  }

  function studentRow(r, rank, isTop) {
    const color = isTop ? "#2E5A1A" : "#8B2222";
    const bg    = isTop ? "#EDF4EA" : "#FAECEC";
    const arrow = isTop ? "▲" : "▼";
    const i = r.idx % 5;
    const slopeAbs = Math.abs(r.slope).toFixed(1);
    const lbl = slopeLabel(r.slope);
    // 首次 vs 最後原始差值作為輔助說明
    const rawColor = r.rawDiff >= 0 ? "#2E5A1A" : "#8B2222";
    const rawSign  = r.rawDiff >= 0 ? "+" : "";
    return `<div style="display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid #F0EAE0">
      <div style="width:24px;height:24px;border-radius:50%;background:${bg};color:${color};font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${rank}</div>
      <div style="width:36px;height:36px;border-radius:50%;background:${RL[i]};color:${RC[i]};font-size:15px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${r.st.name[0]}</div>
      <div style="flex:1;min-width:0">
        <div style="font-size:13px;font-weight:600;color:#1C1A14">${r.st.name}</div>
        <div style="font-size:10px;color:#9E9890;margin-top:2px">${r.firstEx.name} ${r.firstT.toFixed(0)}→${r.lastEx.name} ${r.lastT.toFixed(0)} 分（總差 <span style="color:${rawColor};font-weight:600">${rawSign}${r.rawDiff.toFixed(0)}</span>）</div>
        <div style="font-size:10px;color:#9E9890">${r.dataCount} 次資料 · 趨勢：${lbl}</div>
      </div>
      <div style="text-align:right;flex-shrink:0">
        <div style="font-size:16px;font-weight:700;color:${color};font-family:'DM Mono',monospace">${arrow}${slopeAbs}<span style="font-size:10px;font-weight:400">/次</span></div>
        <div style="font-size:10px;color:#9E9890">回歸斜率</div>
      </div>
    </div>`;
  }

  const top5    = rows.filter(r=>r.slope>0).slice(0,5);
  const bottom5 = [...rows].reverse().filter(r=>r.slope<0).slice(0,5);

  let html = `
    <div style="font-size:11px;color:#9E9890;margin-bottom:10px;padding:8px 10px;background:#FAF7F0;border-radius:6px;border:1px solid #E0DAD0">
      📌 排名依據：<strong>線性回歸斜率</strong>（每次段考平均變化量），反映整段趨勢，不受單次異常影響。括號內為首次到最後一次的原始分差。
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
    <div>
      <div style="font-size:12px;font-weight:600;color:#2E5A1A;letter-spacing:.05em;text-transform:uppercase;margin-bottom:8px;padding-bottom:6px;border-bottom:2px solid #EDF4EA">🌟 進步最多</div>
      ${top5.length ? top5.map((r,i)=>studentRow(r,i+1,true)).join("") : '<div style="font-size:12px;color:#9E9890;padding:8px 0">暫無進步資料</div>'}
    </div>
    <div>
      <div style="font-size:12px;font-weight:600;color:#8B2222;letter-spacing:.05em;text-transform:uppercase;margin-bottom:8px;padding-bottom:6px;border-bottom:2px solid #FAECEC">📉 需要關注</div>
      ${bottom5.length ? bottom5.map((r,i)=>studentRow(r,i+1,false)).join("") : '<div style="font-size:12px;color:#9E9890;padding:8px 0">暫無退步資料</div>'}
    </div>
  </div>`;

  wrap.innerHTML = html;
}

// ══════════════════════════════════════════════════════════
// 功能 2：各科跨學期平均趨勢折線圖
// ══════════════════════════════════════════════════════════
function renderSubjectLongTrend() {
  destroyChart("chart-subject-long");
  const ctx = $("chart-subject-long"); if (!ctx) return;

  const datasets = ACTIVE_SUBJECTS.map((sub, i) => ({
    label: sub,
    data: ACTIVE_EXAMS.map(ex => getExamSubjectAvg(ex.id, sub)),
    borderColor: COLORS[i],
    backgroundColor: COLORS[i] + "18",
    borderWidth: 2,
    pointRadius: 3,
    tension: 0.3,
    spanGaps: true,
    hidden: true,  // 預設隱藏，點擊圖例才顯示
  }));

  // 學期分隔線（垂直線）
  const semStarts = [];
  SEMESTERS.forEach((sem, si) => {
    if (si === 0) return;
    const idx = ACTIVE_EXAMS.findIndex(e => e.semester === sem);
    if (idx > 0) semStarts.push(idx);
  });

  chartInstances["chart-subject-long"] = new Chart(ctx, {
    type: "line",
    data: {
      labels: ACTIVE_EXAMS.map(e => e.name.replace("次段考","").replace("第","").trim()),
      datasets
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: "bottom", labels: { font:{size:11}, boxWidth:12, padding:8 }},
        annotation: {}
      },
      scales: {
        y: { min:0, max:100, grid:{color:"#E2DED6"}, ticks:{font:{size:11}},
             title:{display:true, text:"班級各科平均", font:{size:10}, color:"#9E9890"} },
        x: { grid:{color:"#F5F0E8"}, ticks:{font:{size:9}, maxRotation:45} }
      }
    }
  });
}

// ══════════════════════════════════════════════════════════
// 功能 3：學生分組分析（前段/中段/後段）
// ══════════════════════════════════════════════════════════
function renderGroupAnalysis() {
  destroyChart("chart-group-trend");
  const ctx = $("chart-group-trend"); if (!ctx) return;
  const wrap = $("group-stats-wrap"); if (!wrap) return;

  // ── 每次段考動態計算分組 ──────────────────────────────────
  // 有資料的段考清單
  const activeExams = ACTIVE_EXAMS.filter(ex =>
    S.students.some(st => getFilledCount(getScores(st.id, ex.id)) > 0)
  );
  if (!activeExams.length) {
    wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">尚無資料</div>';
    return;
  }

  // 每次段考：依該次總分將有資料的學生分成三組
  function getExamGroups(examId) {
    const ranked = S.students.map(st => {
      const t = getTotal(getScores(st.id, examId));
      return t !== null ? { st, t } : null;
    }).filter(Boolean).sort((a,b) => b.t - a.t);
    const n = ranked.length;
    if (n < 3) return null;
    const third = Math.floor(n/3);
    return {
      top:    ranked.slice(0, third).map(r=>r.st.id),
      mid:    ranked.slice(third, third*2).map(r=>r.st.id),
      bottom: ranked.slice(third*2).map(r=>r.st.id),
    };
  }

  // 取最新段考的分組做統計卡片
  const latestEx = activeExams[activeExams.length-1];
  const latestGroups = getExamGroups(latestEx.id);
  if (!latestGroups) {
    wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">需要至少 3 位學生才能分組</div>';
    return;
  }

  const groupDef = {
    top:    { label:"前段", color:"#2D5F8A", bg:"#EAF1F8" },
    mid:    { label:"中段", color:"#C4651A", bg:"#FDF0E6" },
    bottom: { label:"後段", color:"#A83232", bg:"#FAECEC" },
  };

  // ── 統計卡片（最新段考分組）──────────────────────────────
  let statsHtml = `<div style="font-size:11px;color:#9E9890;margin-bottom:10px;padding:7px 10px;background:#FAF7F0;border-radius:6px;border:1px solid #E0DAD0">
    📌 依每次段考實際總分動態分組（各組人數盡量相等），組別每次段考可能不同。
  </div>
  <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:16px">`;
  Object.entries(groupDef).forEach(([key, g]) => {
    const ids = latestGroups[key];
    const totals = ids.map(id => getTotal(getScores(id, latestEx.id))).filter(v=>v!==null);
    const gAvg = totals.reduce((a,b)=>a+b,0)/totals.length;
    const gMax = Math.max(...totals), gMin = Math.min(...totals);
    const grpMax = getExamMaxScore(latestEx.id);
    const passLine = grpMax * 0.6;
    const passRate = totals.filter(v=>v>=passLine).length/totals.length*100;
    statsHtml += `<div style="background:${g.bg};border-radius:10px;padding:12px;border:1px solid ${g.color}33">
      <div style="font-size:11px;font-weight:600;color:${g.color};letter-spacing:.05em;margin-bottom:6px">${g.label}（${ids.length}人）<span style="font-weight:400;color:#9E9890"> · ${latestEx.name}</span></div>
      <div style="font-size:20px;font-weight:700;font-family:'DM Mono',monospace;color:${g.color}">${gAvg.toFixed(1)}</div>
      <div style="font-size:10px;color:#9E9890;margin-top:3px">最高 ${gMax.toFixed(0)} / 最低 ${gMin.toFixed(0)}</div>
      <div style="font-size:10px;color:${passRate>=70?"#2E5A1A":passRate<50?"#8B2222":"#C4651A"};margin-top:2px;font-weight:600">及格率 ${passRate.toFixed(0)}%</div>
    </div>`;
  });
  statsHtml += '</div>';

  // ── 追蹤每位學生的組別變化 ───────────────────────────────
  // 計算每位學生在每次段考的組別 (0=前段, 1=中段, 2=後段, null=無資料)
  const groupIndex = { top:0, mid:1, bottom:2 };
  const groupName  = ["前段","中段","後段"];
  const groupColor = ["#2D5F8A","#C4651A","#A83232"];

  const studentTrack = S.students.map(st => {
    const track = activeExams.map(ex => {
      const g = getExamGroups(ex.id);
      if (!g) return null;
      if (g.top.includes(st.id))    return 0;
      if (g.mid.includes(st.id))    return 1;
      if (g.bottom.includes(st.id)) return 2;
      return null; // 無資料
    });
    // 找第一個和最後一個有資料的組別
    const valid = track.filter(v=>v!==null);
    if (valid.length < 2) return null;
    const firstG = valid[0], lastG = valid[valid.length-1];
    const change = lastG - firstG; // 負=升組, 正=降組, 0=不變
    return { st, track, firstG, lastG, change };
  }).filter(Boolean);

  // 升組名單（後→中、中→前、後→前）
  const upgraded   = studentTrack.filter(r=>r.change < 0).sort((a,b)=>a.change-b.change);
  // 降組名單（前→中、中→後、前→後）
  const downgraded = studentTrack.filter(r=>r.change > 0).sort((a,b)=>b.change-a.change);
  // 不變名單
  const stable     = studentTrack.filter(r=>r.change === 0);

  function trackBadge(track) {
    return track.map(g => g===null
      ? `<span style="display:inline-block;width:20px;height:20px;border-radius:50%;background:#F0EAE0;border:1px dashed #C8BA9E;vertical-align:middle;margin:1px"></span>`
      : `<span style="display:inline-block;width:20px;height:20px;border-radius:50%;background:${groupColor[g]};color:#fff;font-size:9px;font-weight:700;line-height:20px;text-align:center;vertical-align:middle;margin:1px">${["前","中","後"][g]}</span>`
    ).join("→");
  }

  let changeHtml = `<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px">`;

  // 升組
  changeHtml += `<div>
    <div style="font-size:12px;font-weight:600;color:#2E5A1A;padding-bottom:6px;border-bottom:2px solid #EDF4EA;margin-bottom:8px">
      ⬆️ 升組（${upgraded.length}人）
    </div>`;
  if (upgraded.length) {
    upgraded.forEach(r => {
      changeHtml += `<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid #F0EAE0">
        <span style="font-size:13px;font-weight:600;color:#1C1A14;min-width:60px">${r.st.name}</span>
        <span style="font-size:11px;color:#9E9890;margin-right:4px">${groupName[r.firstG]}→${groupName[r.lastG]}</span>
        <span>${trackBadge(r.track)}</span>
      </div>`;
    });
  } else {
    changeHtml += '<div style="font-size:12px;color:#9E9890;padding:6px 0">暫無</div>';
  }
  changeHtml += '</div>';

  // 降組
  changeHtml += `<div>
    <div style="font-size:12px;font-weight:600;color:#8B2222;padding-bottom:6px;border-bottom:2px solid #FAECEC;margin-bottom:8px">
      ⬇️ 降組（${downgraded.length}人）
    </div>`;
  if (downgraded.length) {
    downgraded.forEach(r => {
      changeHtml += `<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid #F0EAE0">
        <span style="font-size:13px;font-weight:600;color:#1C1A14;min-width:60px">${r.st.name}</span>
        <span style="font-size:11px;color:#9E9890;margin-right:4px">${groupName[r.firstG]}→${groupName[r.lastG]}</span>
        <span>${trackBadge(r.track)}</span>
      </div>`;
    });
  } else {
    changeHtml += '<div style="font-size:12px;color:#9E9890;padding:6px 0">暫無</div>';
  }
  changeHtml += '</div>';
  changeHtml += '</div>';

  // 穩定名單（摺疊顯示）
  changeHtml += `<div style="font-size:11px;color:#9E9890;margin-bottom:16px">
    ➡️ 組別穩定（${stable.length}人）：${stable.map(r=>`<span style="color:#6B5F4A;font-weight:500">${r.st.name}</span>（${groupName[r.lastG]}）`).join("、")||"—"}
  </div>`;

  wrap.innerHTML = statsHtml + changeHtml;

  // ── 折線圖：依最新分組畫三條走勢線 ──────────────────────
  const datasets = Object.entries(groupDef).map(([key, g]) => {
    const ids = latestGroups[key];
    return {
      label: g.label + `（${ids.length}人）`,
      data: ACTIVE_EXAMS.map(ex => {
        const totals = ids.map(id=>getTotal(getScores(id,ex.id))).filter(v=>v!==null);
        return totals.length >= Math.ceil(ids.length/2) ? totals.reduce((a,b)=>a+b,0)/totals.length : null;
      }),
      borderColor: g.color,
      backgroundColor: g.color + "18",
      borderWidth: 2.5,
      pointRadius: 4,
      tension: 0.3,
      spanGaps: true,
      fill: false,
    };
  });

  chartInstances["chart-group-trend"] = new Chart(ctx, {
    type: "line",
    data: {
      labels: ACTIVE_EXAMS.map(e => e.name.replace("次段考","").replace("第","").trim()),
      datasets
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position:"bottom", labels:{ font:{size:11}, boxWidth:12 }}},
      scales: {
        y: { min:0, max:ACTIVE_SUBJECTS.length*100, grid:{color:"#E2DED6"}, ticks:{font:{size:11}},
             title:{display:true, text:"平均總分", font:{size:10}, color:"#9E9890"} },
        x: { grid:{color:"#F5F0E8"}, ticks:{font:{size:9}, maxRotation:45} }
      }
    }
  });
}

// ══════════════════════════════════════════════════════════
// 功能 4：導師筆記
// ══════════════════════════════════════════════════════════
async function loadNotes() {
  const wrap = $("notes-list-wrap"); if (!wrap) return;
  let notes = {};
  try {
    if (db) {
      const doc = await col("config").doc("teacher_notes").get();
      if (doc.exists) notes = doc.data();
    } else {
      const raw = localStorage.getItem("grade-113-3-notes");
      if (raw) notes = JSON.parse(raw);
    }
  } catch(e) { console.warn("讀取筆記失敗:", e); }

  renderNotesList(notes);
  renderNotesForm(notes);
}

function renderNotesList(notes) {
  const wrap = $("notes-list-wrap"); if (!wrap) return;
  const entries = Object.entries(notes)
    .filter(([k]) => k.startsWith("sem_"))
    .sort(([a],[b]) => a.localeCompare(b));

  if (!entries.length) {
    wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">尚未新增任何筆記</div>';
    return;
  }

  wrap.innerHTML = entries.map(([key, note]) => {
    const sem = key.replace("sem_", "");
    const hasData = ACTIVE_EXAMS.filter(e=>e.semester===sem).some(ex =>
      S.students.some(st => getFilledCount(getScores(st.id, ex.id)) > 0)
    );
    return `<div style="padding:12px 0;border-bottom:1px solid #F0EAE0">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
        <span style="display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:99px;background:${hasData?"#EAF1F8":"#FAF7F0"};color:${hasData?"#2D5F8A":"#9E9890"}">${sem}學期</span>
        <button onclick="deleteNote('${sem}')" style="font-size:10px;color:#9E9890;background:none;border:none;cursor:pointer;padding:2px 6px;border-radius:4px" onmouseover="this.style.color='#8B2222'" onmouseout="this.style.color='#9E9890'">✕ 刪除</button>
      </div>
      <div style="font-size:13px;color:#1C1A14;line-height:1.6;white-space:pre-wrap">${note.text||""}</div>
      <div style="font-size:10px;color:#C8BA9E;margin-top:4px">${note.updatedAt?new Date(note.updatedAt).toLocaleString("zh-TW"):""}</div>
    </div>`;
  }).join("");
}

function renderNotesForm(notes) {
  const semSel = $("note-sem-sel"); if (!semSel) return;
  const cur = semSel.value;
  semSel.innerHTML = SEMESTERS.map(s=>`<option value="${s}">${s}學期</option>`).join("");
  if (cur) semSel.value = cur;
  onNotesSemChange(notes);
}

function onNotesSemChange(notes) {
  const sem = $("note-sem-sel")?.value;
  const ta  = $("note-textarea");
  if (!ta || !sem) return;
  const key = "sem_" + sem;
  ta.value = (notes && notes[key]) ? notes[key].text || "" : "";
}

async function saveNote() {
  const sem  = $("note-sem-sel")?.value;
  const text = $("note-textarea")?.value?.trim();
  const btn  = $("save-note-btn");
  if (!sem) return;

  btn.disabled = true; btn.textContent = "儲存中...";
  try {
    const key = "sem_" + sem;
    const payload = { [key]: { text, updatedAt: new Date().toISOString() } };
    if (db) {
      await col("config").doc("teacher_notes").set(payload, { merge: true });
    }
    // 同步本地
    const raw = localStorage.getItem("grade-113-3-notes");
    const local = raw ? JSON.parse(raw) : {};
    local[key] = payload[key];
    localStorage.setItem("grade-113-3-notes", JSON.stringify(local));

    showToast(`✅ ${sem}學期筆記已儲存`);
    loadNotes();
  } catch(e) { showToast("儲存失敗：" + e.message); }
  finally { btn.disabled = false; btn.textContent = "💾 儲存筆記"; }
}

async function deleteNote(sem) {
  if (!confirm(`確定要刪除 ${sem}學期 的筆記？`)) return;
  const key = "sem_" + sem;
  try {
    if (db) {
      await col("config").doc("teacher_notes").update({ [key]: firebase.firestore.FieldValue.delete() });
    }
    const raw = localStorage.getItem("grade-113-3-notes");
    if (raw) {
      const local = JSON.parse(raw);
      delete local[key];
      localStorage.setItem("grade-113-3-notes", JSON.stringify(local));
    }
    showToast(`🗑️ ${sem}學期筆記已刪除`);
    loadNotes();
  } catch(e) { showToast("刪除失敗：" + e.message); }
}


// ── 不及格名單快速顯示 ────────────────────────────────────────
function showFailList() {
  const examEl = $("input-exam");
  if (!examEl || !examEl.value) { showToast("請先選擇段考"); return; }
  const examId   = examEl.value;
  const examName = examEl.options[examEl.selectedIndex]?.text || examId;
  const maxScore = getExamMaxScore(examId);
  const passLine = maxScore * 0.6;

  const fails = S.students.map(st => {
    const sc = getScores(st.id, examId);
    const total = getTotal(sc);
    const failSubs = ACTIVE_SUBJECTS.filter(sub => {
      const v = sc[sub];
      return v !== undefined && v !== "" && parseFloat(v) < 60;
    });
    return { st, total, failSubs };
  }).filter(x => x.failSubs.length > 0)
    .sort((a,b) => (a.total||9999) - (b.total||9999));

  const modal = $("fail-list-modal");
  const body  = $("fail-list-body");
  if (!modal || !body) { showToast("介面元素缺失"); return; }

  if (!fails.length) {
    body.innerHTML = `<div class="empty-state small"><div class="empty-icon">🎉</div><div class="empty-title">本次段考全班無不及格科目！</div></div>`;
  } else {
    let html = `<div style="margin-bottom:12px;font-size:12px;color:#6B5F4A">共 <strong>${fails.length}</strong> 位學生有不及格科目</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead><tr>
        <th style="background:#FAF7F0;padding:6px 10px;text-align:left;font-size:11px;color:#6B5F4A;border-bottom:1px solid #E0DAD0">座號</th>
        <th style="background:#FAF7F0;padding:6px 10px;text-align:left;font-size:11px;color:#6B5F4A;border-bottom:1px solid #E0DAD0">姓名</th>
        <th style="background:#FAF7F0;padding:6px 10px;text-align:left;font-size:11px;color:#6B5F4A;border-bottom:1px solid #E0DAD0">不及格科目</th>
        <th style="background:#FAF7F0;padding:6px 10px;text-align:center;font-size:11px;color:#6B5F4A;border-bottom:1px solid #E0DAD0">總分</th>
      </tr></thead><tbody>`;
    fails.forEach(({st, total, failSubs}) => {
      const tc = total !== null && total < passLine ? "#8B2222" : "#1C1A14";
      html += `<tr>
        <td style="padding:6px 10px;border-bottom:1px solid #EEE8E0;font-family:monospace">${String(st.number||"").padStart(2,"0")}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #EEE8E0;font-weight:600">${st.name}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #EEE8E0">
          ${failSubs.map(s=>`<span style="display:inline-block;background:#FAECEC;color:#8B2222;border-radius:4px;padding:1px 7px;font-size:11px;margin:1px">${s}</span>`).join(" ")}
        </td>
        <td style="padding:6px 10px;border-bottom:1px solid #EEE8E0;text-align:center;font-family:monospace;font-weight:700;color:${tc}">${total !== null ? total.toFixed(0) : "—"}</td>
      </tr>`;
    });
    html += `</tbody></table>`;
    body.innerHTML = html;
  }
  $("fail-list-title").textContent = `${examName} · 不及格名單`;
  modal.style.display = "flex";
}

function closeFailListModal() {
  const modal = $("fail-list-modal");
  if (modal) modal.style.display = "none";
}

// ── 清除全班所有段考的校排資料 ──────────────────────────────
async function clearAllSchoolRanks() {
  // 第一次確認
  if (!confirm("⚠️ 確定要清除全班所有段考的校排資料？\n\n此動作將刪除所有校排記錄，且無法復原。")) return;
  // 第二次確認（要求輸入）
  const input = prompt("請輸入「確認清除」以執行此操作：");
  if (input !== "確認清除") { showToast("已取消，未清除任何資料"); return; }

  const btn = $("clear-school-rank-btn");
  btn.disabled = true; btn.textContent = "清除中...";
  let count = 0;
  try {
    const batch = db ? db.batch() : null;
    S.students.forEach(st => {
      ACTIVE_EXAMS.forEach(ex => {
        const key = scoreKey(st.id, ex.id);
        if (S.scores[key] && S.scores[key]["校排"] !== undefined) {
          delete S.scores[key]["校排"];
          count++;
          if (batch) batch.update(col("scores").doc(key), {
            "校排": firebase.firestore.FieldValue.delete()
          });
        }
      });
    });
    if (batch) await batch.commit();
    saveLocalData();
    showToast(`✅ 已清除 ${count} 筆校排資料`);
  } catch(e) {
    showToast("清除失敗：" + e.message);
  } finally {
    btn.disabled = false; btn.textContent = "🗑️ 清除全班校排資料";
  }
}

// ── 清除所有成績資料（核彈級）────────────────────────────────
async function clearAllData() {
  // 第一次確認
  if (!confirm("⚠️ 警告：此操作將清除「本次段考所有學生的成績資料」！\n\n包含：所有科目分數、班排、校排。\n學生名單不受影響。\n\n此動作無法復原，確定要繼續嗎？")) return;

  // 選擇清除範圍
  const examEl = $("input-exam");
  const examId = examEl ? examEl.value : null;
  const examName = examEl ? examEl.options[examEl.selectedIndex]?.text : "";

  const scope = prompt(
    `請選擇清除範圍，輸入數字後按確定：\n\n` +
    `1 → 僅清除「${examName}」的成績\n` +
    `2 → 清除「目前學期」所有段考成績\n` +
    `3 → 清除「全部學年」所有成績（最危險）\n\n` +
    `輸入 1、2 或 3：`
  );
  if (!["1","2","3"].includes(scope)) { showToast("已取消，未清除任何資料"); return; }

  // 第二次確認（輸入驗證）
  const scopeLabel = scope === "1" ? `「${examName}」` : scope === "2" ? "目前學期所有段考" : "全部學年所有成績";
  const confirm2 = prompt(`即將清除 ${scopeLabel} 的所有成績。\n\n請輸入「確認清除」以繼續：`);
  if (confirm2 !== "確認清除") { showToast("已取消，未清除任何資料"); return; }

  const btn = $("clear-all-data-btn");
  btn.disabled = true; btn.textContent = "清除中...";

  try {
    // 決定要清除哪些 examId
    let targetExamIds = [];
    if (scope === "1") {
      if (!examId) { showToast("請先選擇段考"); return; }
      targetExamIds = [examId];
    } else if (scope === "2") {
      const semEl = $("input-sem");
      const sem = semEl ? semEl.value : null;
      targetExamIds = ACTIVE_EXAMS.filter(e => e.semester === sem).map(e => e.id);
    } else {
      targetExamIds = ACTIVE_EXAMS.map(e => e.id);
    }

    let count = 0;
    const scoreFields = [...ACTIVE_SUBJECTS, "班排", "校排", "_updatedAt"];

    if (db) {
      // Firestore 每批最多 500 筆，分批處理
      const allKeys = [];
      S.students.forEach(st => {
        targetExamIds.forEach(eid => {
          const key = scoreKey(st.id, eid);
          if (S.scores[key] && Object.keys(S.scores[key]).length > 0) {
            allKeys.push(key);
          }
        });
      });

      // 分批刪除（每批 400 筆）
      for (let i = 0; i < allKeys.length; i += 400) {
        const batch = db.batch();
        allKeys.slice(i, i + 400).forEach(key => {
          batch.delete(col("scores").doc(key));
        });
        await batch.commit();
        count += allKeys.slice(i, i + 400).length;
      }
    }

    // 清除本地 S.scores
    S.students.forEach(st => {
      targetExamIds.forEach(eid => {
        const key = scoreKey(st.id, eid);
        if (S.scores[key]) {
          delete S.scores[key];
          count++;
        }
      });
    });

    clearExamSubjectCountCache();
    saveLocalData();
    renderInputTable();
    showToast(`✅ 已清除 ${targetExamIds.length} 次段考、共 ${count} 筆成績資料`);
  } catch(e) {
    showToast("清除失敗：" + e.message);
    console.error(e);
  } finally {
    btn.disabled = false; btn.textContent = "⚠️ 清除所有成績資料";
  }
}

// ══════════════════════════════════════════════════════════
// 功能 3：預警門檻設定（存在 config 裡）
// ══════════════════════════════════════════════════════════
const DEFAULT_ALERT = {
  scoreThreshold: 60, consecutiveCount: 2, failSubjectCount: 3, schoolRankThreshold: 0,
  subjectThresholds: {}  // { "國文": 60, "英語文": 60, ... } — 各科自訂及格門檻
};

function getAlertConfig() {
  try {
    const raw = localStorage.getItem("grade-alert-config");
    return raw ? { ...DEFAULT_ALERT, ...JSON.parse(raw) } : { ...DEFAULT_ALERT };
  } catch { return { ...DEFAULT_ALERT }; }
}

// 取得某科的及格門檻（優先用個別設定，沒有則用全域 scoreThreshold）
function getSubjectPassLine(sub) {
  const cfg = getAlertConfig();
  return cfg.subjectThresholds?.[sub] ?? cfg.scoreThreshold ?? 60;
}

function saveAlertConfig() {
  const cfg = {
    scoreThreshold:      parseInt($("alert-score-threshold")?.value)       || 60,
    consecutiveCount:    parseInt($("alert-consecutive-count")?.value)     || 2,
    failSubjectCount:    parseInt($("alert-fail-subject-count")?.value)    || 3,
    schoolRankThreshold: parseInt($("alert-school-rank-threshold")?.value) || 0,
    subjectThresholds:   {}
  };
  // 各科個別門檻
  ACTIVE_SUBJECTS.forEach(sub => {
    const el = document.getElementById("subthresh-" + sub.replace(/\s/g,"_"));
    if (el && el.value !== "" && parseInt(el.value) !== cfg.scoreThreshold) {
      cfg.subjectThresholds[sub] = parseInt(el.value);
    }
  });
  localStorage.setItem("grade-alert-config", JSON.stringify(cfg));
  if (db) col("config").doc("alert_config").set(cfg).catch(()=>{});
  showToast("✅ 預警門檻已儲存");
  // 重新渲染需關注名單
  if (S.activePage === "overview") renderOverview();
}

async function loadAlertConfig() {
  let cfg = { ...DEFAULT_ALERT };
  try {
    if (db) {
      const doc = await col("config").doc("alert_config").get();
      if (doc.exists) cfg = { ...DEFAULT_ALERT, ...doc.data() };
    }
    const raw = localStorage.getItem("grade-alert-config");
    if (raw) cfg = { ...cfg, ...JSON.parse(raw) };
  } catch {}
  // 寫回全域 UI
  if ($("alert-score-threshold"))        $("alert-score-threshold").value        = cfg.scoreThreshold;
  if ($("alert-consecutive-count"))      $("alert-consecutive-count").value      = cfg.consecutiveCount;
  if ($("alert-fail-subject-count"))     $("alert-fail-subject-count").value     = cfg.failSubjectCount;
  if ($("alert-school-rank-threshold"))  $("alert-school-rank-threshold").value  = cfg.schoolRankThreshold||0;
  // 產生各科門檻欄位
  const grid = $("subject-threshold-grid");
  if (grid) {
    grid.innerHTML = ACTIVE_SUBJECTS.map(sub => {
      const key = sub.replace(/\s/g,"_");
      const val = cfg.subjectThresholds?.[sub] ?? "";
      return `<div>
        <label style="font-size:12px;color:#6B5F4A;display:block;margin-bottom:4px">${sub}</label>
        <input type="number" id="subthresh-${key}" min="0" max="100" value="${val}"
          placeholder="${cfg.scoreThreshold}"
          style="width:100%;padding:8px 10px;border:1px solid #C8BA9E;border-radius:6px;font-size:13px">
      </div>`;
    }).join("");
  }
  return cfg;
}

// ══════════════════════════════════════════════════════════
// 功能 6：學生名單搜尋
// ══════════════════════════════════════════════════════════
function filterStudentList(query) {
  const q = query.trim().toLowerCase();
  const rows = document.querySelectorAll("#student-list-table tbody tr");
  rows.forEach(row => {
    const text = row.textContent.toLowerCase();
    row.style.display = (!q || text.includes(q)) ? "" : "none";
  });
}

// ══════════════════════════════════════════════════════════
// 功能 7：段考前提醒清單
// ══════════════════════════════════════════════════════════
function renderAlertList() {
  const wrap = $("alert-list-wrap"); if (!wrap) return;
  const cfg = getAlertConfig();

  // 找最新有資料的段考
  const latestEx = [...ACTIVE_EXAMS].reverse().find(ex =>
    S.students.some(st => getFilledCount(getScores(st.id, ex.id)) > 0)
  );
  if (!latestEx) { wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">尚無段考資料</div>'; return; }

  const alerts = S.students.map(st => {
    const sc = getScores(st.id, latestEx.id);
    // 低於門檻的科目
    const failSubs = ACTIVE_SUBJECTS.filter(sub => {
      const v = sc[sub];
      return v!==undefined && v!=="" && parseFloat(v) < cfg.scoreThreshold;
    });
    // 連續低於門檻
    const consecSubs = ACTIVE_SUBJECTS.filter(sub => {
      const vals = ACTIVE_EXAMS.map(ex => {
        const v = getScores(st.id, ex.id)[sub];
        return v!==undefined && v!=="" ? parseFloat(v) : null;
      }).filter(v=>v!==null);
      const recent = vals.slice(-cfg.consecutiveCount);
      return recent.length >= cfg.consecutiveCount && recent.every(v => v < cfg.scoreThreshold);
    });
    const sr = sc["校排"] ? parseInt(sc["校排"]) : null;
    const srAlert = cfg.schoolRankThreshold > 0 && sr !== null && sr > cfg.schoolRankThreshold;
    if (failSubs.length < cfg.failSubjectCount && consecSubs.length === 0 && !srAlert) return null;
    return { st, failSubs, consecSubs, sr, srAlert };
  }).filter(Boolean);

  if (!alerts.length) {
    wrap.innerHTML = `<div style="font-size:13px;color:#2E5A1A;padding:12px 0">🎉 目前沒有需要特別關注的學生！（門檻：${cfg.scoreThreshold}分以下 ${cfg.failSubjectCount}科以上）</div>`;
    return;
  }

  const COLORS = ["#2D5F8A","#5B8A4A","#C4651A","#6B4FA0","#A83232"];
  const LIGHTS = ["#EAF1F8","#EDF4EA","#FDF0E6","#F2EEF8","#FAECEC"];

  const alertRows = alerts.map(({st, failSubs, consecSubs, sr, srAlert}, i) => {
    return `<div style="display:flex;align-items:flex-start;gap:12px;padding:12px 0;border-bottom:1px solid #F0EAE0;cursor:pointer" onclick="jumpToAnalysis('${st.id}')">
      <div style="width:38px;height:38px;border-radius:50%;background:${LIGHTS[i%5]};color:${COLORS[i%5]};font-size:15px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${st.name[0]}</div>
      <div style="flex:1">
        <div style="font-size:14px;font-weight:600;color:#1C1A14;margin-bottom:4px">${st.name} <span style="font-size:11px;color:#9E9890">座號 ${st.number||"—"}</span></div>
        ${failSubs.length >= cfg.failSubjectCount ? `<div style="font-size:12px;color:#8B2222;margin-bottom:3px">⚠️ 不及格 ${failSubs.length} 科：${failSubs.join("、")}</div>` : ""}
        ${srAlert ? `<div style="font-size:12px;color:#C4651A">🏫 校排第 ${sr} 名（門檻${cfg.schoolRankThreshold}名）</div>` : ""}
        ${consecSubs.length ? `<div style="font-size:12px;color:#C4651A">🔁 連續不及格：${consecSubs.join("、")}</div>` : ""}
      </div>
      <div style="font-size:11px;color:#2D5F8A;flex-shrink:0">點擊查看 →</div>
    </div>`;
  }).join("");

  wrap.innerHTML = `
    <div style="font-size:12px;color:#6B5F4A;margin-bottom:10px">依據最新段考（${latestEx.name}）資料，門檻：單科低於 ${cfg.scoreThreshold} 分且 ${cfg.failSubjectCount} 科以上，或連續 ${cfg.consecutiveCount} 次低於門檻</div>
    ${alertRows}
  `;
}

// ══════════════════════════════════════════════════════════
// 功能 8：家長通知單列印
// ══════════════════════════════════════════════════════════
function renderParentNotice() {
  const studentId = S.reportStudentId;
  const wrap = $("parent-notice-wrap"); if (!wrap) return;
  if (!studentId) { wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">請先選擇學生</div>'; return; }

  const st = S.students.find(s=>s.id===studentId); if (!st) return;
  const allScores = ACTIVE_EXAMS.map(ex => getScores(studentId, ex.id));

  // 跟報告頁選單連動，若有選段考就用，否則找最新有資料的
  const selectedExamId = $("report-exam")?.value;
  const selectedIdx = selectedExamId ? ACTIVE_EXAMS.findIndex(e=>e.id===selectedExamId) : -1;
  const latestIdx = selectedIdx >= 0 && getFilledCount(allScores[selectedIdx]) > 0
    ? selectedIdx
    : [...Array(ACTIVE_EXAMS.length).keys()].reverse().find(i => getFilledCount(allScores[i]) > 0);
  if (latestIdx === undefined) { wrap.innerHTML = '<div style="font-size:12px;color:#9E9890">尚無成績資料</div>'; return; }

  // 從 S.teacherComments 讀取評語（與報告頁共用同一份）
  const examIdForComment = selectedExamId || "all";
  const savedComment = (S.teacherComments[studentId] || {})[examIdForComment]
    || $("report-teacher-comment")?.innerText?.trim() || "";

  const sc    = allScores[latestIdx];
  const ex    = ACTIVE_EXAMS[latestIdx];
  const total = getTotal(sc);
  const avg   = getAvg(sc);

  // 與上次比較（找此次之前最近有資料的）
  const prevIdx = [...Array(latestIdx).keys()].reverse().find(i => getFilledCount(allScores[i]) > 0);
  const prevTotal = prevIdx !== undefined ? getTotal(allScores[prevIdx]) : null;
  const diffTotal = (total !== null && prevTotal !== null) ? total - prevTotal : null;

  // 不及格科目
  const failSubs = ACTIVE_SUBJECTS.filter(sub => {
    const v = sc[sub];
    return v !== undefined && v !== "" && parseFloat(v) < 60;
  });

  wrap.innerHTML = `
    <div class="parent-notice no-screen" style="border:2px solid #1C1A14;border-radius:8px;padding:24px;max-width:600px;font-family:'Noto Sans TC',sans-serif;background:#fff;page-break-before:always;page-break-inside:avoid;break-before:page;break-inside:avoid">
      <div style="text-align:center;border-bottom:2px solid #1C1A14;padding-bottom:12px;margin-bottom:16px">
        <div style="font-size:11px;color:#6B5F4A;letter-spacing:.1em">${getClassName()} · ${getClassYear()}</div>
        <div style="font-size:20px;font-weight:900;margin:4px 0">段考成績通知單</div>
        <div style="font-size:12px;color:#6B5F4A">${ex.name} · 列印日期：${new Date().toLocaleDateString("zh-TW")}</div>
      </div>
      <div style="display:flex;gap:20px;margin-bottom:16px">
        <div><span style="font-size:12px;color:#6B5F4A">姓名：</span><strong>${st.name}</strong></div>
        <div><span style="font-size:12px;color:#6B5F4A">座號：</span><strong>${st.number||"—"}</strong></div>
        <div><span style="font-size:12px;color:#6B5F4A">班排名：</span><strong>${sc["班排"]||"—"}</strong> 名</div>
        <div><span style="font-size:12px;color:#6B5F4A">校排名：</span><strong>${sc["校排"]||"—"}</strong> 名</div>
      </div>
      <table style="width:100%;border-collapse:collapse;margin-bottom:16px;font-size:13px;page-break-inside:avoid;break-inside:avoid">
        <thead><tr style="background:#F5F0E8">
          ${ACTIVE_SUBJECTS.map(s=>`<th style="padding:6px 8px;text-align:center;border:1px solid #C8BA9E;font-size:11px">${s}</th>`).join("")}
          <th style="padding:6px 8px;text-align:center;border:1px solid #C8BA9E;background:#E8E0D0">總分</th>
          <th style="padding:6px 8px;text-align:center;border:1px solid #C8BA9E;background:#E8E0D0">平均</th>
        </tr></thead>
        <tbody><tr>
          ${ACTIVE_SUBJECTS.map(sub=>{
            const v=sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null;
            const color=v===null?"#C8BA9E":v<60?"#8B2222":"#1C1A14";
            return `<td style="padding:8px;text-align:center;border:1px solid #C8BA9E;font-family:monospace;font-weight:600;color:${color}">${v!==null?v.toFixed(0):"—"}</td>`;
          }).join("")}
          <td style="padding:8px;text-align:center;border:1px solid #C8BA9E;font-weight:700;font-size:15px;background:#F5F0E8">${total!==null?total.toFixed(0):"—"}</td>
          <td style="padding:8px;text-align:center;border:1px solid #C8BA9E;background:#F5F0E8">${avg!==null?avg.toFixed(1):"—"}</td>
        </tr></tbody>
      </table>
      <div style="display:flex;gap:12px;margin-bottom:16px;font-size:12px">
        ${diffTotal!==null?`<div style="padding:6px 12px;border-radius:6px;background:${diffTotal>=0?"#EDF4EA":"#FAECEC"};color:${diffTotal>=0?"#2E5A1A":"#8B2222"};font-weight:600">
          與上次相比：${diffTotal>=0?"▲":"▼"}${Math.abs(diffTotal).toFixed(0)} 分
        </div>`:""}
        ${failSubs.length?`<div style="padding:6px 12px;border-radius:6px;background:#FAECEC;color:#8B2222;font-weight:600">
          ⚠️ 不及格：${failSubs.join("、")}
        </div>`:""}
      </div>
      <div style="margin-bottom:20px">
        <div style="font-size:12px;font-weight:600;color:#6B5F4A;margin-bottom:6px">導師評語
          <span style="font-size:10px;font-weight:400;color:#9E9890;margin-left:6px">（可直接點擊輸入）</span>
        </div>
        <div id="parent-notice-teacher-comment"
             contenteditable="true"
             spellcheck="false"
             data-placeholder="點此輸入導師評語..."
             style="border:1px solid #C8BA9E;border-radius:4px;padding:10px;min-height:60px;font-size:13px;color:#1C1A14;outline:none;cursor:text;transition:border .15s"
             onfocus="this.style.borderColor='#8B7355';this.style.background='#FDFBF7'"
             onblur="this.style.borderColor='#C8BA9E';this.style.background=''"
        ></div>
      </div>
      <div style="display:flex;justify-content:space-between;font-size:13px;padding-top:12px;border-top:1px solid #C8BA9E">
        <div>家長簽名：_______________</div>
        <div>日期：_______________</div>
        <div>導師簽名：_______________</div>
      </div>
    </div>
  `;

  // 還原評語，並綁定雙向同步 + debounce 存 Firebase
  const commentEl = $("parent-notice-teacher-comment");
  if (commentEl && savedComment) commentEl.innerText = savedComment;
  if (commentEl) {
    let saveTimer;
    commentEl.oninput = () => {
      const txt = commentEl.innerText.trim();
      if (!S.teacherComments[studentId]) S.teacherComments[studentId] = {};
      S.teacherComments[studentId][examIdForComment] = txt;
      const rc = $("report-teacher-comment");
      if (rc) rc.innerText = commentEl.innerText;
      clearTimeout(saveTimer);
      saveTimer = setTimeout(() => saveTeacherComments(), 2000);
    };
  }
}


// ── 全班批次列印家長通知單 ────────────────────────────────────
function batchPrintAllNotices() {
  if (!S.students.length) { showToast("尚無學生資料"); return; }

  // 找目前選取的段考（跟隨 report-exam 選單）
  const examEl = $("report-exam");
  const examId = examEl?.value;
  const examName = examEl?.options[examEl.selectedIndex]?.text || "";
  if (!examId) { showToast("請先在上方選擇段考"); return; }

  // 只列印有成績的學生
  const validStudents = S.students.filter(st =>
    getFilledCount(getScores(st.id, examId)) > 0
  );
  if (!validStudents.length) { showToast("本次段考尚無任何學生有成績資料"); return; }

  const hideRank = $("report-hide-rank")?.checked || false;
  const maxScore = getExamMaxScore(examId);
  const passLine = maxScore * 0.6;

  // 產生每位學生的通知單 HTML
  const pages = validStudents.map(st => {
    const sc       = getScores(st.id, examId);
    const total    = getTotal(sc);
    const avg      = getAvg(sc);
    const ex       = ACTIVE_EXAMS.find(e=>e.id===examId);

    // 與上次比較
    const stIdx    = ACTIVE_EXAMS.findIndex(e=>e.id===examId);
    const allScores = ACTIVE_EXAMS.map(ex2 => getScores(st.id, ex2.id));
    const prevIdx  = [...Array(stIdx).keys()].reverse().find(i => getFilledCount(allScores[i]) > 0);
    const prevTotal = prevIdx !== undefined ? getTotal(allScores[prevIdx]) : null;
    const diffTotal = (total!==null&&prevTotal!==null) ? total-prevTotal : null;

    // 不及格科目
    const failSubs = ACTIVE_SUBJECTS.filter(sub => {
      const v = sc[sub];
      return v!==undefined && v!=="" && parseFloat(v) < 60;
    });

    const savedComment = (S.teacherComments[st.id]||{})[examId]||"";

    return `<div style="border:2px solid #1C1A14;border-radius:8px;padding:24px;font-family:'Noto Sans TC',sans-serif;background:#fff;page-break-inside:avoid;break-inside:avoid;page-break-after:always">
      <div style="text-align:center;border-bottom:2px solid #1C1A14;padding-bottom:12px;margin-bottom:16px">
        <div style="font-size:11px;color:#6B5F4A;letter-spacing:.1em">${getClassName()} · ${getClassYear()}</div>
        <div style="font-size:20px;font-weight:900;margin:4px 0">段考成績通知單</div>
        <div style="font-size:12px;color:#6B5F4A">${ex?.name||examId} · 列印日期：${new Date().toLocaleDateString("zh-TW")}</div>
      </div>
      <div style="display:flex;gap:20px;margin-bottom:16px">
        <div><span style="font-size:12px;color:#6B5F4A">姓名：</span><strong>${st.name}</strong></div>
        <div><span style="font-size:12px;color:#6B5F4A">座號：</span><strong>${st.number||"—"}</strong></div>
        ${!hideRank?`<div><span style="font-size:12px;color:#6B5F4A">班排名：</span><strong>${sc["班排"]||"—"}</strong> 名</div>
        <div><span style="font-size:12px;color:#6B5F4A">校排名：</span><strong>${sc["校排"]||"—"}</strong> 名</div>`:""}
      </div>
      <table style="width:100%;border-collapse:collapse;margin-bottom:16px;font-size:13px">
        <thead><tr style="background:#F5F0E8">
          ${ACTIVE_SUBJECTS.map(s=>`<th style="padding:6px 8px;text-align:center;border:1px solid #C8BA9E;font-size:11px">${s}</th>`).join("")}
          <th style="padding:6px 8px;text-align:center;border:1px solid #C8BA9E;background:#E8E0D0">總分</th>
          <th style="padding:6px 8px;text-align:center;border:1px solid #C8BA9E;background:#E8E0D0">平均</th>
        </tr></thead>
        <tbody><tr>
          ${ACTIVE_SUBJECTS.map(sub=>{
            const v=sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null;
            const color=v===null?"#C8BA9E":v<60?"#8B2222":"#1C1A14";
            return `<td style="padding:8px;text-align:center;border:1px solid #C8BA9E;font-family:monospace;font-weight:600;color:${color}">${v!==null?v.toFixed(0):"—"}</td>`;
          }).join("")}
          <td style="padding:8px;text-align:center;border:1px solid #C8BA9E;font-weight:700;font-size:15px;background:#F5F0E8">${total!==null?total.toFixed(0):"—"}</td>
          <td style="padding:8px;text-align:center;border:1px solid #C8BA9E;background:#F5F0E8">${avg!==null?avg.toFixed(1):"—"}</td>
        </tr></tbody>
      </table>
      <div style="display:flex;gap:12px;margin-bottom:16px;font-size:12px">
        ${diffTotal!==null?`<div style="padding:6px 12px;border-radius:6px;background:${diffTotal>=0?"#EDF4EA":"#FAECEC"};color:${diffTotal>=0?"#2E5A1A":"#8B2222"};font-weight:600">與上次相比：${diffTotal>=0?"▲":"▼"}${Math.abs(diffTotal).toFixed(0)} 分</div>`:""}
        ${failSubs.length?`<div style="padding:6px 12px;border-radius:6px;background:#FAECEC;color:#8B2222;font-weight:600">⚠️ 不及格：${failSubs.join("、")}</div>`:""}
      </div>
      <div style="margin-bottom:20px">
        <div style="font-size:12px;font-weight:600;color:#6B5F4A;margin-bottom:6px">導師評語</div>
        <div style="border:1px solid #C8BA9E;border-radius:4px;padding:10px;min-height:60px;font-size:13px">${savedComment||""}</div>
      </div>
      <div style="display:flex;justify-content:space-between;font-size:13px;padding-top:12px;border-top:1px solid #C8BA9E">
        <div>家長簽名：_______________</div>
        <div>日期：_______________</div>
        <div>導師簽名：_______________</div>
      </div>
    </div>`;
  });

  // 組合並列印，隱藏其他區塊
  const wrap = $("parent-notice-wrap");
  const origHtml = wrap.innerHTML;
  wrap.innerHTML = pages.join("");

  const pageReport = document.getElementById("page-report");
  const hiddenEls = [];
  if (pageReport) {
    Array.from(pageReport.children).forEach(el => {
      if (!el.contains(wrap)) {
        el.style.display = "none";
        hiddenEls.push(el);
      }
    });
    const noticeSection = document.getElementById("parent-notice-section");
    if (noticeSection) {
      Array.from(noticeSection.children).forEach(el => {
        if (el.id !== "parent-notice-wrap") {
          el.style.display = "none";
          hiddenEls.push(el);
        }
      });
    }
  }

  showToast(`⏳ 準備列印 ${pages.length} 位學生的通知單...`);
  setTimeout(() => {
    window.print();
    setTimeout(() => {
      hiddenEls.forEach(el => el.style.display = "");
      wrap.innerHTML = origHtml;
    }, 800);
  }, 300);
}

function printParentNotice() {
  renderParentNotice();
  setTimeout(() => {
    // 隱藏報告卡、頁面標題、通知單的標題和選單列，只保留通知單本體
    const pageReport = document.getElementById("page-report");
    const noticeWrap = document.getElementById("parent-notice-wrap");
    // 把 page-report 內除了 parent-notice-wrap 以外的所有子元素都隱藏
    const hiddenEls = [];
    if (pageReport) {
      Array.from(pageReport.children).forEach(el => {
        if (!el.contains(noticeWrap)) {
          el.style.display = "none";
          hiddenEls.push(el);
        }
      });
      // 隱藏 parent-notice-section 裡面除了 parent-notice-wrap 的其他元素
      const noticeSection = document.getElementById("parent-notice-section");
      if (noticeSection) {
        Array.from(noticeSection.children).forEach(el => {
          if (el.id !== "parent-notice-wrap") {
            el.style.display = "none";
            hiddenEls.push(el);
          }
        });
      }
    }
    window.print();
    setTimeout(() => {
      hiddenEls.forEach(el => el.style.display = "");
    }, 800);
  }, 300);
}

// ══════════════════════════════════════════════════════════
// 校排分析工具函式
// ══════════════════════════════════════════════════════════

// 取得全班某次段考的校排陣列（有填的）
function getExamSchoolRanks(examId) {
  return S.students.map(st => {
    const r = getScores(st.id, examId)["校排"];
    return r ? parseInt(r) : null;
  });
}

// 取得全班某次段考的平均校排
function getExamAvgSchoolRank(examId) {
  const ranks = getExamSchoolRanks(examId).filter(v=>v!==null);
  return ranks.length ? ranks.reduce((a,b)=>a+b,0)/ranks.length : null;
}

// ── 功能 2：全班校排分布圖 ───────────────────────────────
function renderSchoolRankDist(examId) {
  const wrap = $("school-rank-dist-wrap"); if (!wrap) return;
  const ranks = getExamSchoolRanks(examId).filter(v=>v!==null);
  if (!ranks.length) { wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">尚無校排資料</div>'; return; }

  const maxRank = Math.max(...ranks);
  // 分成5個區間（依校排百分位）
  const bandSize = Math.ceil(maxRank / 5);
  const bands = [
    { label:`前20%（1-${bandSize}名）`,       min:1,           max:bandSize,   color:"#1A5C2A" },
    { label:`20-40%`,                         min:bandSize+1,  max:bandSize*2, color:"#5B8A4A" },
    { label:`40-60%`,                         min:bandSize*2+1,max:bandSize*3, color:"#2D5F8A" },
    { label:`60-80%`,                         min:bandSize*3+1,max:bandSize*4, color:"#C4651A" },
    { label:`後20%（${bandSize*4+1}名以後）`, min:bandSize*4+1,max:Infinity,   color:"#A83232" },
  ];
  const counts = bands.map(b => ranks.filter(r=>r>=b.min&&r<=b.max).length);
  const avgRank = ranks.reduce((a,b)=>a+b,0)/ranks.length;

  let html = `<div style="font-size:12px;color:#6B5F4A;margin-bottom:10px">
    班級平均校排：<strong>${avgRank.toFixed(1)}</strong> 名　最佳：<strong>${Math.min(...ranks)}</strong> 名　最差：<strong>${Math.max(...ranks)}</strong> 名
  </div>`;
  html += bands.map((b,i) => `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
      <span style="font-size:11px;color:#6B5F4A;min-width:160px">${b.label}</span>
      <div style="flex:1;height:20px;background:#F0EAE0;border-radius:4px;overflow:hidden">
        <div style="height:100%;width:${counts[i]?counts[i]/ranks.length*100:0}%;background:${b.color};border-radius:4px;transition:width .4s"></div>
      </div>
      <span style="font-size:13px;font-weight:700;color:${b.color};min-width:28px;text-align:right">${counts[i]}</span>
      <span style="font-size:11px;color:#9E9890">人</span>
    </div>`).join("");
  wrap.innerHTML = html;
}

// ── 功能 4：班級平均校排走勢 ─────────────────────────────
function renderSchoolRankTrend() {
  destroyChart("chart-school-rank-trend");
  const ctx = $("chart-school-rank-trend"); if (!ctx) return;
  const avgRanks = ACTIVE_EXAMS.map(ex => getExamAvgSchoolRank(ex.id));
  const hasData = avgRanks.some(v=>v!==null);
  if (!hasData) return;

  chartInstances["chart-school-rank-trend"] = new Chart(ctx, {
    type:"line",
    data:{ labels: ACTIVE_EXAMS.map(e=>e.name.replace("次段考","").replace("第","")),
      datasets:[{ label:"平均校排", data:avgRanks,
        borderColor:"#C4651A", backgroundColor:"#C4651A18",
        borderWidth:2.5, pointRadius:4, fill:true, tension:0.3, spanGaps:true
      }]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false},
        tooltip:{ callbacks:{ label: ctx => `平均校排：第 ${ctx.raw?.toFixed(1)||"—"} 名` } }
      },
      scales:{
        y:{ reverse:true, grid:{color:"#E2DED6"}, ticks:{font:{size:11}},
          title:{display:true,text:"校排名（越小越好）",font:{size:10},color:"#9E9890"} },
        x:{ grid:{color:"#F0EAE0"}, ticks:{font:{size:9},maxRotation:45} }
      }
    }
  });
}

// ── 功能 6：班內強但校內弱警示 ───────────────────────────
function renderClassStrongSchoolWeak(examId) {
  const wrap = $("class-strong-school-weak-wrap"); if (!wrap) return;
  const ranks = S.students.map(st => {
    const sc = getScores(st.id, examId);
    return { st, classRank: sc["班排"]?parseInt(sc["班排"]):null, schoolRank: sc["校排"]?parseInt(sc["校排"]):null };
  }).filter(r=>r.classRank!==null&&r.schoolRank!==null);

  if (!ranks.length) { wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">尚無班排+校排資料</div>'; return; }

  const n = S.students.length;
  // 班排前33%但校排後50%
  const warnings = ranks.filter(r => r.classRank <= Math.ceil(n*0.33) && r.schoolRank > ranks.reduce((a,b)=>a+(b.schoolRank||0),0)/ranks.length );
  if (!warnings.length) {
    wrap.innerHTML = '<div style="font-size:13px;color:#2E5A1A;padding:8px">✅ 目前沒有班排與校排落差過大的學生</div>';
    return;
  }
  const COLORS=["#2D5F8A","#5B8A4A","#C4651A","#6B4FA0","#A83232"];
  const LIGHTS=["#EAF1F8","#EDF4EA","#FDF0E6","#F2EEF8","#FAECEC"];
  wrap.innerHTML = `
    <div style="font-size:12px;color:#6B5F4A;margin-bottom:8px">班排前 33% 但校排低於班級平均，需特別留意</div>
    ${warnings.map(({st,classRank,schoolRank},i)=>`
    <div style="display:flex;align-items:center;gap:12px;padding:8px 0;border-bottom:1px solid #F0EAE0;cursor:pointer" onclick="jumpToAnalysis('${st.id}')">
      <div style="width:36px;height:36px;border-radius:50%;background:${LIGHTS[i%5]};color:${COLORS[i%5]};font-size:14px;font-weight:700;display:flex;align-items:center;justify-content:center">${st.name[0]}</div>
      <div style="flex:1">
        <div style="font-size:13px;font-weight:600">${st.name}</div>
        <div style="font-size:11px;color:#9E9890">班排 第${classRank}名 · 校排 第${schoolRank}名</div>
      </div>
      <div style="font-size:11px;color:#C4651A;font-weight:600">班內強 ⚠️ 校內弱</div>
    </div>`).join("")}`;
}

// ══════════════════════════════════════════════════════════
// 功能 5：班級健康指數
// ══════════════════════════════════════════════════════════
function calcClassHealthIndex(examId, gradeInfo) {
  let totals, avg, passRate, stdDev, maxScore, trendScore;

  if (gradeInfo) {
    // ── 年級總結模式：跨所有段考的全班總分 ──────────────────
    const gradeExams = ACTIVE_EXAMS.filter(e => gradeInfo.semesters.includes(e.semester));
    // 每位學生在這個年級的平均總分
    const studentAvgs = S.students.map(st => {
      const valid = gradeExams.map(ex => getTotal(getScores(st.id, ex.id))).filter(v=>v!==null);
      return valid.length ? valid.reduce((a,b)=>a+b,0)/valid.length : null;
    }).filter(v=>v!==null);
    if (studentAvgs.length < 3) return null;

    maxScore = getExamMaxScore(gradeExams[0].id) || 900;
    const passLine = maxScore * 0.6;
    const n = studentAvgs.length;
    avg      = studentAvgs.reduce((a,b)=>a+b,0)/n;
    passRate = studentAvgs.filter(v=>v>=passLine).length/n*100;
    stdDev   = Math.sqrt(studentAvgs.reduce((a,v)=>a+Math.pow(v-avg,2),0)/n);
    totals   = studentAvgs;

    // 趨勢：比較第一個學期平均 vs 最後一個學期平均
    const firstSemExams = gradeExams.filter(e=>e.semester===gradeInfo.semesters[0]);
    const lastSemExams  = gradeExams.filter(e=>e.semester===gradeInfo.semesters[gradeInfo.semesters.length-1]);
    const firstAvgs = S.students.map(st=>{
      const v=firstSemExams.map(ex=>getTotal(getScores(st.id,ex.id))).filter(v=>v!==null);
      return v.length?v.reduce((a,b)=>a+b,0)/v.length:null;
    }).filter(v=>v!==null);
    const lastAvgs = S.students.map(st=>{
      const v=lastSemExams.map(ex=>getTotal(getScores(st.id,ex.id))).filter(v=>v!==null);
      return v.length?v.reduce((a,b)=>a+b,0)/v.length:null;
    }).filter(v=>v!==null);
    if (firstAvgs.length>=3 && lastAvgs.length>=3) {
      const firstMean = firstAvgs.reduce((a,b)=>a+b,0)/firstAvgs.length;
      const lastMean  = lastAvgs.reduce((a,b)=>a+b,0)/lastAvgs.length;
      trendScore = Math.min(100, Math.max(0, 50 + (lastMean-firstMean)*2));
    } else { trendScore = 50; }

  } else {
    // ── 單次段考模式（原本邏輯）─────────────────────────────
    totals = S.students.map(st => getTotal(getScores(st.id, examId))).filter(v=>v!==null);
    if (totals.length < 3) return null;
    maxScore = getExamMaxScore(examId);
    const passLine = maxScore * 0.6;
    const n  = totals.length;
    avg      = totals.reduce((a,b)=>a+b,0)/n;
    passRate = totals.filter(v=>v>=passLine).length/n*100;
    stdDev   = Math.sqrt(totals.reduce((a,v)=>a+Math.pow(v-avg,2),0)/n);
    const exIdx = ACTIVE_EXAMS.findIndex(e=>e.id===examId);
    trendScore = 50;
    if (exIdx > 0) {
      const prevTotals = S.students.map(st=>getTotal(getScores(st.id,ACTIVE_EXAMS[exIdx-1].id))).filter(v=>v!==null);
      if (prevTotals.length>=3) {
        const prevAvg = prevTotals.reduce((a,b)=>a+b,0)/prevTotals.length;
        trendScore = Math.min(100, Math.max(0, 50+(avg-prevAvg)*2));
      }
    }
  }

  const maxT = Math.max(...totals), minT = Math.min(...totals);
  const range = maxT - minT;
  const avgScore    = Math.min(100, (avg/maxScore)*100);
  const passScore   = passRate;
  const spreadScore = Math.max(0, 100-(range/maxScore)*100);
  const stdScore    = Math.max(0, 100-(stdDev/maxScore*2)*100);
  const health      = Math.round(avgScore*0.4 + passScore*0.3 + trendScore*0.2 + stdScore*0.1);

  return {
    health: Math.min(100, Math.max(0, health)),
    avg, passRate, stdDev, range, trendScore,
    components: { avgScore, passScore, trendScore, stdScore }
  };
}

function renderHealthIndex(examId, gradeInfo) {
  const wrap = $("health-index-wrap"); if (!wrap) return;
  const result = calcClassHealthIndex(examId, gradeInfo);
  if (!result) { wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">需要至少 3 位學生的資料</div>'; return; }

  const modeLabel = gradeInfo ? gradeInfo.label : "本次段考";
  const { health, avg, passRate, stdDev, trendScore } = result;

  // 顏色與評語
  let color, label, desc;
  if (health >= 85)      { color="#1A5C2A"; label="非常健康 🌟"; desc="班級整體表現優秀，進步穩定，分布集中。"; }
  else if (health >= 70) { color="#2E7A3A"; label="健康 ✅";     desc="班級表現良好，大多數學生達到水準。"; }
  else if (health >= 55) { color="#2D5F8A"; label="普通 ℹ️";    desc="班級表現中等，部分學生需要加強。"; }
  else if (health >= 40) { color="#C4651A"; label="需關注 ⚠️";  desc="班級整體偏弱，建議加強輔導。"; }
  else                   { color="#8B2222"; label="需緊急關注 🚨"; desc="班級表現偏低，建議立即安排補救教學。"; }

  const ring = `
    <div style="position:relative;width:100px;height:100px;flex-shrink:0">
      <svg viewBox="0 0 100 100" width="100" height="100">
        <circle cx="50" cy="50" r="44" fill="none" stroke="#F0EAE0" stroke-width="10"/>
        <circle cx="50" cy="50" r="44" fill="none" stroke="${color}" stroke-width="10"
          stroke-dasharray="${health * 2.764} 276.4"
          stroke-dashoffset="69.1" stroke-linecap="round" transform="rotate(-90 50 50)"/>
      </svg>
      <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center">
        <div style="font-size:22px;font-weight:900;color:${color};line-height:1">${health}</div>
        <div style="font-size:9px;color:#9E9890">/ 100</div>
      </div>
    </div>`;

  const bars = [
    { label:"平均分佔比", val: result.components.avgScore, desc:`${avg.toFixed(1)}分` },
    { label:"及格率",     val: passRate,                   desc:`${passRate.toFixed(1)}%` },
    { label:"進步趨勢",   val: trendScore,                 desc:trendScore>50?"進步中":trendScore<50?"退步中":"持平" },
    { label:"分布集中度", val: result.components.stdScore, desc:`標差${stdDev.toFixed(1)}` },
  ];

  const barsHtml = bars.map(b=>{
    const bc = b.val>=70?"#5B8A4A":b.val>=50?"#2D5F8A":"#A83232";
    return `<div style="margin-bottom:8px">
      <div style="display:flex;justify-content:space-between;font-size:11px;color:#6B5F4A;margin-bottom:3px">
        <span>${b.label}</span><span style="font-weight:600;color:${bc}">${b.desc}</span>
      </div>
      <div style="height:6px;background:#F0EAE0;border-radius:3px;overflow:hidden">
        <div style="height:100%;width:${b.val.toFixed(1)}%;background:${bc};border-radius:3px;transition:width .6s"></div>
      </div>
    </div>`;
  }).join("");

  wrap.innerHTML = `
    <div style="font-size:11px;color:#9E9890;background:#FAF7F0;border:1px solid #E0DAD0;border-radius:6px;padding:7px 12px;margin-bottom:10px;line-height:1.7">
      📐 計算方式：<strong>平均分佔比</strong> ×40% ＋ <strong>及格率</strong> ×30% ＋ <strong>進步趨勢</strong> ×20% ＋ <strong>分布集中度</strong> ×10%
      <span style="margin-left:8px;color:#C8BA9E">（資料範圍：${modeLabel}）</span>
    </div>
    <div style="display:flex;gap:20px;align-items:center;flex-wrap:wrap">
      ${ring}
      <div style="flex:1;min-width:180px">
        <div style="font-size:16px;font-weight:700;color:${color};margin-bottom:4px">${label}</div>
        <div style="font-size:12px;color:#6B5F4A;margin-bottom:12px">${desc}</div>
        ${barsHtml}
      </div>
    </div>`;
}

// ══════════════════════════════════════════════════════════
// 功能 4：成績預測
// ══════════════════════════════════════════════════════════
function linearRegression(ys) {
  // 只用有資料的點做線性回歸
  const pts = ys.map((y,i)=>y!==null?{x:i,y}:null).filter(Boolean);
  if (pts.length < 2) return null;
  const n  = pts.length;
  const sx = pts.reduce((a,p)=>a+p.x,0);
  const sy = pts.reduce((a,p)=>a+p.y,0);
  const sx2= pts.reduce((a,p)=>a+p.x*p.x,0);
  const sxy= pts.reduce((a,p)=>a+p.x*p.y,0);
  const denom = n*sx2 - sx*sx;
  if (denom===0) return null;
  const slope = (n*sxy - sx*sy) / denom;
  const intercept = (sy - slope*sx) / n;
  return { slope, intercept };
}

function renderPrediction(examId) {
  const wrap = $("prediction-wrap"); if (!wrap) return;
  if (!S.students.length) { wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">尚無學生資料</div>'; return; }

  const exIdx = ACTIVE_EXAMS.findIndex(e=>e.id===examId);
  if (exIdx < 0) return;

  // 計算每位學生的預測下一次總分
  const predictions = S.students.map(st => {
    const totals = ACTIVE_EXAMS.map(ex => getTotal(getScores(st.id, ex.id)));
    const reg = linearRegression(totals.slice(0, exIdx+1));
    if (!reg) return null;
    const nextIdx = exIdx + 1;
    if (nextIdx >= ACTIVE_EXAMS.length) return null;
    const predicted = Math.round(reg.slope * nextIdx + reg.intercept);
    const maxScore  = getExamMaxScore(ACTIVE_EXAMS[Math.min(nextIdx, ACTIVE_EXAMS.length-1)].id) || getExamMaxScore(examId);
    const clamped   = Math.max(0, Math.min(maxScore, predicted));
    const current   = totals[exIdx];
    const trend     = reg.slope; // 每次段考的平均變化
    return { st, predicted: clamped, current, trend, nextEx: ACTIVE_EXAMS[nextIdx] };
  }).filter(Boolean);

  if (!predictions.length) {
    wrap.innerHTML = '<div style="font-size:12px;color:#9E9890;padding:8px">資料不足（需至少 2 次段考）</div>';
    return;
  }

  // 依預測分數排序
  predictions.sort((a,b) => b.predicted - a.predicted);
  const nextExName = predictions[0]?.nextEx?.name || "下一次段考";
  const COLORS = ["#2D5F8A","#5B8A4A","#C4651A","#6B4FA0","#A83232"];
  const LIGHTS = ["#EAF1F8","#EDF4EA","#FDF0E6","#F2EEF8","#FAECEC"];

  // 預測需關注（預測分低於及格線，或預測比本次退步超過20分，或長期斜率 < -10）
  const maxScore = getExamMaxScore(examId);
  const passLine = maxScore * 0.6;
  const atRisk   = predictions.filter(p => {
    const diff = p.current!==null ? p.predicted - p.current : null;
    return p.predicted < passLine || (diff!==null && diff < -20) || p.trend < -10;
  });

  // 圖例說明
  const legendHtml = `
    <div style="display:flex;flex-wrap:wrap;gap:10px;margin-bottom:14px;font-size:11px;color:#6B5F4A">
      <div style="display:flex;align-items:center;gap:6px;background:#FAF7F0;border:1px solid #E0DAD0;border-radius:8px;padding:5px 10px">
        <span style="font-weight:700">預測趨勢</span>
        <span>＝ 預測下次 − 本次總分（短期預測方向）</span>
      </div>
      <div style="display:flex;align-items:center;gap:6px;background:#FAF7F0;border:1px solid #E0DAD0;border-radius:8px;padding:5px 10px">
        <span style="font-weight:700">學習走勢</span>
        <span>＝ 歷次成績回歸斜率（長期整體趨勢）</span>
      </div>
    </div>`;

  let html = `
    <div style="font-size:12px;color:#6B5F4A;margin-bottom:10px">
      基於歷次段考趨勢預測「${nextExName}」的可能表現（線性回歸估算，僅供參考）
    </div>
    ${legendHtml}`;

  if (atRisk.length) {
    html += `<div style="background:#FFF8F0;border:1px solid #F0C090;border-radius:8px;padding:12px;margin-bottom:14px">
      <div style="font-size:12px;font-weight:600;color:#C4651A;margin-bottom:8px">⚠️ 預測可能需要關注（${atRisk.length}人）</div>
      ${atRisk.map(p=>`<span style="display:inline-block;background:#FAECEC;color:#8B2222;font-size:11px;padding:2px 8px;border-radius:99px;margin:2px;font-weight:600">${p.st.name} 預測${p.predicted}</span>`).join("")}
    </div>`;
  }

  html += `<div class="table-wrap"><table>
    <thead><tr>
      <th>學生</th>
      <th style="text-align:center">本次總分</th>
      <th style="text-align:center">預測下次</th>
      <th style="text-align:center" title="預測下次 − 本次總分">預測趨勢<br><span style="font-size:10px;font-weight:400;color:#9E9890">短期方向</span></th>
      <th style="text-align:center" title="歷次成績線性回歸斜率">學習走勢<br><span style="font-size:10px;font-weight:400;color:#9E9890">長期趨勢</span></th>
      <th style="min-width:90px">預測分佈</th>
    </tr></thead>
    <tbody>
    ${predictions.map(({st,predicted,current,trend},i)=>{
      const diff = current!==null ? predicted - current : null;

      // 短期：預測 vs 本次
      const shortColor = diff===null?"#6B5F4A":diff>5?"#2E5A1A":diff<-5?"#8B2222":"#6B5F4A";
      const shortLabel = diff===null?"—"
        :diff>20?`<span style="color:#1A5C2A">快速進步↑</span>`
        :diff>5 ?`<span style="color:#2E5A1A">進步↑</span>`
        :diff>-5?`<span style="color:#6B5F4A">持平→</span>`
        :diff>-20?`<span style="color:#8B2222">退步↓</span>`
        :`<span style="color:#6B0A0A">快速退步↓</span>`;
      const shortSub = diff!==null
        ? `<div style="font-size:10px;color:#9E9890;margin-top:2px">${diff>0?"+":""}${diff.toFixed(0)} 分</div>`
        : "";

      // 長期：回歸斜率
      const longColor = trend>5?"#2E5A1A":trend<-5?"#8B2222":"#6B5F4A";
      const longLabel = trend>10?`<span style="color:#1A5C2A">長期上升↑</span>`
        :trend>5 ?`<span style="color:#2E5A1A">緩步上升↑</span>`
        :trend>-5?`<span style="color:#6B5F4A">長期持平→</span>`
        :trend>-10?`<span style="color:#8B2222">緩步下滑↓</span>`
        :`<span style="color:#6B0A0A">持續下滑↓</span>`;
      const longSub = `<div style="font-size:10px;color:#9E9890;margin-top:2px">每次 ${trend>0?"+":""}${trend.toFixed(1)} 分</div>`;

      // 值得特別注意：短期預測進步但長期走勢下滑（或相反）
      const conflicted = (diff!==null && diff>5 && trend<-5) || (diff!==null && diff<-5 && trend>5);
      const rowBg = conflicted ? "background:#FFFBF0" : "";

      const predictedColor = predicted>=passLine?"#2E5A1A":"#8B2222";
      const barW = Math.min(100, (predicted/maxScore)*100);
      return `<tr style="${rowBg}">
        <td style="font-weight:500">${st.name}${conflicted?` <span title="短期預測與長期走勢方向不一致，需留意" style="font-size:10px;color:#C4651A;cursor:default">⚡</span>`:""}
        </td>
        <td style="text-align:center;font-family:'DM Mono',monospace;font-weight:600">${current!==null?current.toFixed(0):"—"}</td>
        <td style="text-align:center;font-family:'DM Mono',monospace;font-weight:700;font-size:15px;color:${predictedColor}">${predicted}</td>
        <td style="text-align:center;font-size:12px;font-weight:600">${shortLabel}${shortSub}</td>
        <td style="text-align:center;font-size:12px;font-weight:600">${longLabel}${longSub}</td>
        <td style="padding:8px 10px">
          <div style="height:6px;background:#F0EAE0;border-radius:3px;overflow:hidden">
            <div style="height:100%;width:${barW}%;background:${predictedColor};border-radius:3px"></div>
          </div>
        </td>
      </tr>`;
    }).join("")}
    </tbody>
  </table></div>
  <div style="font-size:11px;color:#9E9890;margin-top:8px">⚡ 短期預測與長期走勢方向不一致，建議進一步觀察</div>`;

  wrap.innerHTML = html;
}

// ── 匯出 Excel 成績單 ────────────────────────────────────
function exportExcel() {
  if (typeof XLSX === "undefined") { showToast("⚠️ Excel 套件尚未載入"); return; }
  if (!S.students.length) { showToast("尚無學生資料"); return; }

  const wb = XLSX.utils.book_new();

  // ── 每個有資料的段考建一張 Sheet ────────────────────────
  ACTIVE_EXAMS.forEach(ex => {
    const hasData = S.students.some(st => getFilledCount(getScores(st.id, ex.id)) > 0);
    if (!hasData) return;

    const headers = ["座號","姓名",...ACTIVE_SUBJECTS,"總分","平均","班排","校排"];
    const rows = [headers];
    S.students.forEach(st => {
      const sc = getScores(st.id, ex.id);
      const total = getTotal(sc);
      const avgVal = getAvg(sc);
      const row = [
        st.number||"", st.name,
        ...ACTIVE_SUBJECTS.map(sub => sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):""),
        total!==null?parseFloat(total.toFixed(0)):"",
        avgVal!==null?parseFloat(avgVal.toFixed(1)):"",
        sc["班排"]||"", sc["校排"]||""
      ];
      rows.push(row);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    // 欄寬
    ws["!cols"] = [6,12,...ACTIVE_SUBJECTS.map(()=>8),8,8,6,6].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws, ex.name);
  });

  // ── 總覽 Sheet ───────────────────────────────────────────
  const overviewHeaders = ["座號","姓名",...ACTIVE_EXAMS.map(e=>e.name+" 總分"),...ACTIVE_EXAMS.map(e=>e.name+" 班排"),...ACTIVE_EXAMS.map(e=>e.name+" 校排"),"累積最高","累積最低"];
  const overviewRows = [overviewHeaders];
  S.students.forEach(st => {
    const totals = ACTIVE_EXAMS.map(ex => getTotal(getScores(st.id, ex.id)));
    const ranks  = ACTIVE_EXAMS.map(ex => getScores(st.id, ex.id)["班排"]||"");
    const sranks = ACTIVE_EXAMS.map(ex => getScores(st.id, ex.id)["校排"]||"");
    const valid  = totals.filter(v=>v!==null);
    overviewRows.push([
      st.number||"", st.name,
      ...totals.map(v=>v!==null?parseFloat(v.toFixed(0)):""),
      ...ranks, ...sranks,
      valid.length?Math.max(...valid):"",
      valid.length?Math.min(...valid):""
    ]);
  });
  const wsOverview = XLSX.utils.aoa_to_sheet(overviewRows);
  XLSX.utils.book_append_sheet(wb, wsOverview, "總覽");

  const date = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `113-3班成績單_${date}.xlsx`);
  showToast("✅ Excel 成績單已匯出！");
}

// （診斷工具已移除）



// ── 學期綜合報告 HTML 產生器（供批次列印用，不寫入 DOM）────────
function buildSemesterSummaryHTML(studentId, st, sem, hideRank) {
  const semExams  = ACTIVE_EXAMS.filter(e => e.semester === sem);
  const semScores = semExams.map(ex => getScores(studentId, ex.id));
  const semTotals = semScores.map(sc => getTotal(sc));
  const validPairs = semExams.map((ex,i)=>({ex,sc:semScores[i],t:semTotals[i]})).filter(x=>x.t!==null);
  if (!validPairs.length) return null;

  const firstT = validPairs[0].t, lastT = validPairs[validPairs.length-1].t;
  const diff   = validPairs.length >= 2 ? lastT - firstT : null;
  const maxT   = Math.max(...validPairs.map(p=>p.t));
  const minT   = Math.min(...validPairs.map(p=>p.t));
  const avgT   = validPairs.map(p=>p.t).reduce((a,b)=>a+b,0)/validPairs.length;

  const subAvgs = ACTIVE_SUBJECTS.map(sub => {
    const vals = semScores.map(sc=>sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null).filter(v=>v!==null);
    const clsVals = semExams.flatMap(ex=>S.students.map(st2=>{const v=getScores(st2.id,ex.id)[sub];return v!==undefined&&v!==""?parseFloat(v):null;})).filter(v=>v!==null);
    const myAvg  = vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null;
    const clsAvg = clsVals.length?clsVals.reduce((a,b)=>a+b,0)/clsVals.length:null;
    return { sub, myAvg, clsAvg, count: vals.length };
  }).filter(s=>s.myAvg!==null);

  const bestSub  = subAvgs.length?subAvgs.reduce((a,b)=>a.myAvg>b.myAvg?a:b):null;
  const failSubs = subAvgs.filter(s=>s.myAvg<60);
  const aboveAvgSubs = subAvgs.filter(s=>s.clsAvg!==null&&s.myAvg>=s.clsAvg+5);
  const latestRank = validPairs[validPairs.length-1].sc["班排"]||null;
  const latestSchoolRank = validPairs[validPairs.length-1].sc["校排"]||null;

  const diffHtml = diff===null?"":diff>0
    ? `<span style="color:#2E5A1A;font-weight:700">▲${diff.toFixed(0)} 分進步</span>`
    : diff<0?`<span style="color:#8B2222;font-weight:700">▼${Math.abs(diff).toFixed(0)} 分退步</span>`
    : `<span style="color:#6B5F4A">持平</span>`;

  let compareRows = "";
  semExams.forEach((ex, i) => {
    const sc = semScores[i], t = semTotals[i];
    const prevT = i>0?semTotals[i-1]:null;
    const td = (t!==null&&prevT!==null)?t-prevT:null;
    const tdHtml = td===null?"":td>0?`<span style="color:#2E5A1A;font-size:11px"> ▲${td.toFixed(0)}</span>`:`<span style="color:#8B2222;font-size:11px"> ▼${Math.abs(td).toFixed(0)}</span>`;
    if (getFilledCount(sc)===0) return;
    compareRows += `<tr><td style="font-size:12px;font-weight:600;white-space:nowrap">${ex.name}</td>`;
    ACTIVE_SUBJECTS.forEach(sub => {
      const v = sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null;
      const c = v===null?"#C8BA9E":v>=80?"#2E5A1A":v<60?"#8B2222":"#1C1A14";
      compareRows += `<td style="text-align:center;font-family:monospace;font-size:12px;font-weight:600;color:${c}">${v!==null?v.toFixed(0):"—"}</td>`;
    });
    const tc = t===null?"#C8BA9E":t>=ACTIVE_SUBJECTS.length*80?"#2E5A1A":t<ACTIVE_SUBJECTS.length*60?"#8B2222":"#2D5F8A";
    compareRows += `<td style="text-align:center;font-family:monospace;font-weight:700;font-size:13px;color:${tc}">${t!==null?t.toFixed(0):"—"}${tdHtml}</td>`;
    if (!hideRank) compareRows += `<td style="text-align:center;font-size:12px;color:#6B5F4A">${sc["班排"]||"—"}</td><td style="text-align:center;font-size:12px;color:#6B4FA0">${sc["校排"]||"—"}</td>`;
    compareRows += `</tr>`;
  });

  let avgRow = `<tr style="background:#FAF7F0;border-top:2px solid #C8BA9E"><td style="font-size:11px;font-weight:700;color:#6B5F4A">學期平均</td>`;
  ACTIVE_SUBJECTS.forEach(sub => {
    const a = subAvgs.find(s=>s.sub===sub);
    const c = !a?"#C8BA9E":a.myAvg>=80?"#2E5A1A":a.myAvg<60?"#8B2222":"#1C1A14";
    avgRow += `<td style="text-align:center;font-family:monospace;font-size:12px;font-weight:700;color:${c}">${a?a.myAvg.toFixed(1):"—"}</td>`;
  });
  avgRow += `<td style="text-align:center;font-family:monospace;font-weight:700;font-size:13px;color:#2D5F8A">${avgT.toFixed(1)}</td>`;
  if (!hideRank) avgRow += `<td colspan="2"></td>`;
  avgRow += `</tr>`;

  const savedComment = (S.teacherComments[studentId]||{})[sem+"_summary"]||"";

  return `<div class="report-paper">
    <div class="report-header">
      <div>
        <div class="report-school">${getClassName()} ${sem}學期綜合報告 · ${getClassYear()}</div>
        <div class="report-name">${st.name}</div>
        <div class="report-meta">座號：${st.number||"—"} ／ 列印日期：${new Date().toLocaleDateString("zh-TW")}</div>
      </div>
      <div style="text-align:right">
        <div style="font-size:32px;font-weight:900;font-family:monospace;color:#1C1A14;line-height:1">${avgT.toFixed(0)}</div>
        <div style="font-size:11px;color:#8B7355">學期平均總分</div>
      </div>
    </div>
    <div style="background:linear-gradient(135deg,#1C3A5E,#2D5F8A);color:#EAF1F8;border-radius:6px;padding:10px 14px;margin-bottom:14px;font-size:12px;line-height:2">
      📈 ${sem}學期共 ${validPairs.length} 次段考　${diff!==null?diffHtml+"（"+firstT.toFixed(0)+"→"+lastT.toFixed(0)+"分）":""}
      　　最高 <strong>${maxT.toFixed(0)}</strong> 分　最低 <strong>${minT.toFixed(0)}</strong> 分
      ${bestSub?`<br>💪 最強科目：<strong>${bestSub.sub}</strong>（學期均 ${bestSub.myAvg.toFixed(1)} 分）`:""}
      ${failSubs.length?`<br>⚠️ 學期平均不及格：<strong>${failSubs.map(s=>s.sub+"("+s.myAvg.toFixed(1)+")").join("、")}</strong>`:""}
      ${aboveAvgSubs.length?`<br>📊 高於班平均：<strong>${aboveAvgSubs.map(s=>s.sub).join("、")}</strong>`:""}
      ${!hideRank&&latestRank?`<br>🏅 最新班排：<strong>${latestRank}</strong> 名${latestSchoolRank?" ／ 校排 <strong>"+latestSchoolRank+"</strong> 名":""}`:``}
    </div>
    <div class="report-section" style="margin-bottom:14px">
      <div class="report-exam-title">📊 ${sem}學期三次段考成績對比</div>
      <div class="table-wrap"><table class="report-table">
        <thead><tr>
          <th>段考</th>
          ${ACTIVE_SUBJECTS.map(s=>`<th style="text-align:center;white-space:nowrap">${s}</th>`).join("")}
          <th style="text-align:center">總分</th>
          ${!hideRank?`<th style="text-align:center">班排</th><th style="text-align:center">校排</th>`:""}
        </tr></thead>
        <tbody>${compareRows}${avgRow}</tbody>
      </table></div>
    </div>
    <div class="report-section" style="margin-bottom:14px">
      <div class="report-exam-title">📐 各科學期平均 vs 班級平均</div>
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px">
        ${subAvgs.map(({sub,myAvg,clsAvg})=>{
          const diff2=clsAvg!==null?myAvg-clsAvg:null;
          const bc=myAvg>=80?"#2E5A1A":myAvg<60?"#8B2222":"#2D5F8A";
          const dc=diff2===null?"#9E9890":diff2>=0?"#2E5A1A":"#8B2222";
          return `<div style="background:#FAF7F0;border:1px solid #E0DAD0;border-radius:6px;padding:8px 10px">
            <div style="font-size:11px;font-weight:700;color:#6B5F4A;margin-bottom:4px">${sub}</div>
            <div style="font-family:monospace;font-size:18px;font-weight:900;color:${bc}">${myAvg.toFixed(1)}</div>
            ${clsAvg!==null?`<div style="font-size:10px;color:${dc}">${diff2>=0?"▲":"▼"}${Math.abs(diff2).toFixed(1)} vs 班平均(${clsAvg.toFixed(1)})</div>`:""}
          </div>`;
        }).join("")}
      </div>
    </div>
    <div class="report-comment">
      <div class="report-comment-title">學期綜合評語</div>
      <div class="report-comment-box">${savedComment||"&nbsp;"}</div>
    </div>
    <div class="report-sign">
      <div>導師簽名：_______________</div>
      <div>家長簽名：_______________</div>
      <div>日期：_______________</div>
    </div>
  </div>`;
}

// ── 學期綜合報告 ──────────────────────────────────────────────
function renderSemesterSummaryReport(studentId, st, sem, hideRank) {
  const wrap = $("report-content");
  const semExams  = ACTIVE_EXAMS.filter(e => e.semester === sem);
  const allScores = ACTIVE_EXAMS.map(ex => getScores(studentId, ex.id));
  const semScores = semExams.map(ex => getScores(studentId, ex.id));
  const semTotals = semScores.map(sc => getTotal(sc));
  const validPairs = semExams.map((ex,i)=>({ex,sc:semScores[i],t:semTotals[i]})).filter(x=>x.t!==null);

  if (!validPairs.length) {
    wrap.innerHTML = `<div class="empty-state"><div class="empty-icon">📋</div><div class="empty-title">${sem}學期尚無成績資料</div></div>`;
    return;
  }

  const firstT = validPairs[0].t, lastT = validPairs[validPairs.length-1].t;
  const diff   = validPairs.length >= 2 ? lastT - firstT : null;
  const maxT   = Math.max(...validPairs.map(p=>p.t));
  const minT   = Math.min(...validPairs.map(p=>p.t));
  const avgT   = validPairs.map(p=>p.t).reduce((a,b)=>a+b,0)/validPairs.length;

  // 各科學期平均
  const subAvgs = ACTIVE_SUBJECTS.map(sub => {
    const vals = semScores.map(sc=>sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null).filter(v=>v!==null);
    const clsVals = semExams.flatMap(ex=>S.students.map(st2=>{const v=getScores(st2.id,ex.id)[sub];return v!==undefined&&v!==""?parseFloat(v):null;})).filter(v=>v!==null);
    const myAvg  = vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null;
    const clsAvg = clsVals.length?clsVals.reduce((a,b)=>a+b,0)/clsVals.length:null;
    return { sub, myAvg, clsAvg, count: vals.length };
  }).filter(s=>s.myAvg!==null);

  const bestSub = subAvgs.length?subAvgs.reduce((a,b)=>a.myAvg>b.myAvg?a:b):null;
  const weakSub = subAvgs.length?subAvgs.reduce((a,b)=>a.myAvg<b.myAvg?a:b):null;
  const failSubs = subAvgs.filter(s=>s.myAvg<60);
  const aboveAvgSubs = subAvgs.filter(s=>s.clsAvg!==null&&s.myAvg>=s.clsAvg+5);
  const belowAvgSubs = subAvgs.filter(s=>s.clsAvg!==null&&s.myAvg<=s.clsAvg-5);

  // 從全部段考找前學期末排名（用於趨勢）
  const prevSemExams = ACTIVE_EXAMS.filter(e=>{
    const si = SEMESTERS.indexOf(sem);
    return si>0 && e.semester===SEMESTERS[si-1];
  });
  const prevLastSc = prevSemExams.length?getScores(studentId,prevSemExams[prevSemExams.length-1].id):null;
  const prevRank = prevLastSc?.["班排"]||null;
  const latestRank = validPairs[validPairs.length-1].sc["班排"]||null;
  const latestSchoolRank = validPairs[validPairs.length-1].sc["校排"]||null;

  // 存評語
  const examIdForComment = sem + "_summary";
  const savedComment = (S.teacherComments[studentId]||{})[examIdForComment]||"";

  const diffHtml = diff===null?"":diff>0
    ? `<span style="color:#2E5A1A;font-weight:700">▲${diff.toFixed(0)} 分進步</span>`
    : diff<0?`<span style="color:#8B2222;font-weight:700">▼${Math.abs(diff).toFixed(0)} 分退步</span>`
    : `<span style="color:#6B5F4A">持平</span>`;

  // 三次段考對比表
  let compareRows = "";
  semExams.forEach((ex, i) => {
    const sc = semScores[i], t = semTotals[i];
    const prevIdx = i > 0 ? i-1 : null;
    const prevT = prevIdx!==null?semTotals[prevIdx]:null;
    const td = (t!==null&&prevT!==null)?t-prevT:null;
    const tdHtml = td===null?"":td>0?`<span style="color:#2E5A1A;font-size:11px"> ▲${td.toFixed(0)}</span>`:`<span style="color:#8B2222;font-size:11px"> ▼${Math.abs(td).toFixed(0)}</span>`;
    if (getFilledCount(sc)===0) return;
    compareRows += `<tr>
      <td style="font-size:12px;font-weight:600;white-space:nowrap">${ex.name}</td>`;
    ACTIVE_SUBJECTS.forEach(sub => {
      const v = sc[sub]!==undefined&&sc[sub]!==""?parseFloat(sc[sub]):null;
      const c = v===null?"#C8BA9E":v>=80?"#2E5A1A":v<60?"#8B2222":"#1C1A14";
      compareRows += `<td style="text-align:center;font-family:'DM Mono',monospace;font-size:12px;font-weight:600;color:${c}">${v!==null?v.toFixed(0):"—"}</td>`;
    });
    const tc = t===null?"#C8BA9E":t>=ACTIVE_SUBJECTS.length*80?"#2E5A1A":t<ACTIVE_SUBJECTS.length*60?"#8B2222":"#2D5F8A";
    compareRows += `<td style="text-align:center;font-family:'DM Mono',monospace;font-weight:700;font-size:13px;color:${tc}">${t!==null?t.toFixed(0):"—"}${tdHtml}</td>`;
    if (!hideRank) compareRows += `<td style="text-align:center;font-size:12px;color:#6B5F4A">${sc["班排"]||"—"}</td><td style="text-align:center;font-size:12px;color:#6B4FA0">${sc["校排"]||"—"}</td>`;
    compareRows += `</tr>`;
  });

  // 班级平均列
  let avgRow = `<tr style="background:#FAF7F0;border-top:2px solid #C8BA9E"><td style="font-size:11px;font-weight:700;color:#6B5F4A">學期平均</td>`;
  ACTIVE_SUBJECTS.forEach(sub => {
    const a = subAvgs.find(s=>s.sub===sub);
    const c = !a?"#C8BA9E":a.myAvg>=80?"#2E5A1A":a.myAvg<60?"#8B2222":"#1C1A14";
    avgRow += `<td style="text-align:center;font-family:'DM Mono',monospace;font-size:12px;font-weight:700;color:${c}">${a?a.myAvg.toFixed(1):"—"}</td>`;
  });
  avgRow += `<td style="text-align:center;font-family:'DM Mono',monospace;font-weight:700;font-size:13px;color:#2D5F8A">${avgT.toFixed(1)}</td>`;
  if (!hideRank) avgRow += `<td colspan="2"></td>`;
  avgRow += `</tr>`;

  wrap.innerHTML = `
    <div class="report-paper">
      <div class="report-header">
        <div>
          <div class="report-school">${getClassName()} ${sem}學期綜合報告 · ${getClassYear()}</div>
          <div class="report-name">${st.name}</div>
          <div class="report-meta">座號：${st.number||"—"} ／ 列印日期：${new Date().toLocaleDateString("zh-TW")}</div>
        </div>
        <div style="text-align:right">
          <div style="font-size:32px;font-weight:900;font-family:monospace;color:#1C1A14;line-height:1">${avgT.toFixed(0)}</div>
          <div style="font-size:11px;color:#8B7355">學期平均總分</div>
        </div>
      </div>

      <!-- 學期摘要 -->
      <div style="background:linear-gradient(135deg,#1C3A5E,#2D5F8A);color:#EAF1F8;border-radius:6px;padding:10px 14px;margin-bottom:14px;font-size:12px;line-height:2">
        📈 ${sem}學期共 ${validPairs.length} 次段考　${diff!==null?diffHtml+"（"+firstT.toFixed(0)+"→"+lastT.toFixed(0)+"分）":""}
        　　最高 <strong>${maxT.toFixed(0)}</strong> 分　最低 <strong>${minT.toFixed(0)}</strong> 分
        ${bestSub?`<br>💪 最強科目：<strong>${bestSub.sub}</strong>（學期均 ${bestSub.myAvg.toFixed(1)} 分）`:""}
        ${failSubs.length?`<br>⚠️ 學期平均不及格：<strong>${failSubs.map(s=>s.sub+"("+s.myAvg.toFixed(1)+")").join("、")}</strong>`:""}
        ${aboveAvgSubs.length?`<br>📊 高於班平均：<strong>${aboveAvgSubs.map(s=>s.sub).join("、")}</strong>`:""}
        ${!hideRank&&latestRank?`<br>🏅 最新班排：<strong>${latestRank}</strong> 名${!hideRank&&latestSchoolRank?" ／ 校排 <strong>"+latestSchoolRank+"</strong> 名":""}`:``}
      </div>

      <!-- 三次段考對比表 -->
      <div class="report-section" style="margin-bottom:14px">
        <div class="report-exam-title">📊 ${sem}學期三次段考成績對比</div>
        <div class="table-wrap"><table class="report-table">
          <thead><tr>
            <th>段考</th>
            ${ACTIVE_SUBJECTS.map(s=>`<th style="text-align:center;white-space:nowrap">${s}</th>`).join("")}
            <th style="text-align:center">總分</th>
            ${!hideRank?`<th style="text-align:center">班排</th><th style="text-align:center">校排</th>`:""}
          </tr></thead>
          <tbody>${compareRows}${avgRow}</tbody>
        </table></div>
      </div>

      <!-- 各科學期平均與班級對比 -->
      <div class="report-section" style="margin-bottom:14px">
        <div class="report-exam-title">📐 各科學期平均 vs 班級平均</div>
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px">
          ${subAvgs.map(({sub,myAvg,clsAvg})=>{
            const diff2 = clsAvg!==null?myAvg-clsAvg:null;
            const bc = myAvg>=80?"#2E5A1A":myAvg<60?"#8B2222":"#2D5F8A";
            const dc = diff2===null?"#9E9890":diff2>=0?"#2E5A1A":"#8B2222";
            return `<div style="background:#FAF7F0;border:1px solid #E0DAD0;border-radius:6px;padding:8px 10px">
              <div style="font-size:11px;font-weight:700;color:#6B5F4A;margin-bottom:4px">${sub}</div>
              <div style="font-family:monospace;font-size:18px;font-weight:900;color:${bc}">${myAvg.toFixed(1)}</div>
              ${clsAvg!==null?`<div style="font-size:10px;color:${dc}">${diff2>=0?"▲":"▼"}${Math.abs(diff2).toFixed(1)} vs 班平均(${clsAvg.toFixed(1)})</div>`:""}
            </div>`;
          }).join("")}
        </div>
      </div>

      <!-- 導師評語 -->
      <div class="report-comment">
        <div class="report-comment-title">學期綜合評語
          <span style="font-size:10px;font-weight:400;color:#9E9890;margin-left:6px">（可直接點擊輸入，自動儲存）</span>
        </div>
        <div id="report-teacher-comment"
             class="report-comment-box"
             contenteditable="true"
             spellcheck="false"
             style="color:#1C1A14;outline:none;cursor:text"
             data-placeholder="點此輸入學期綜合評語..."
        ></div>
      </div>
      <div class="report-sign">
        <div>導師簽名：_______________</div>
        <div>家長簽名：_______________</div>
        <div>日期：_______________</div>
      </div>
    </div>`;

  // 載入評語
  const commentEl = $("report-teacher-comment");
  if (commentEl && savedComment) commentEl.innerText = savedComment;
  if (commentEl) {
    let saveTimer;
    commentEl.oninput = () => {
      const txt = commentEl.innerText.trim();
      if (!S.teacherComments[studentId]) S.teacherComments[studentId] = {};
      S.teacherComments[studentId][examIdForComment] = txt;
      const pnc = $("parent-notice-teacher-comment");
      if (pnc) pnc.innerText = commentEl.innerText;
      clearTimeout(saveTimer);
      saveTimer = setTimeout(() => saveTeacherComments(), 2000);
    };
  }
}